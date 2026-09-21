"""Tests for Word Companion Add-in bridge and live commenting integration."""

import json
import urllib.request

import pytest
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from word.bridge import WordBridge, WordSession
from word.certs import get_or_create_dev_certificate


def test_dev_certificate_generation(tmp_path):
    cert_path, key_path = get_or_create_dev_certificate(cert_dir=tmp_path)
    assert cert_path.is_file()
    assert key_path.is_file()
    assert cert_path.stat().st_size > 500
    assert key_path.stat().st_size > 500
    assert (key_path.stat().st_mode & 0o777) == 0o600

    # Re-calling returns the same existing certificate
    c2, k2 = get_or_create_dev_certificate(cert_dir=tmp_path)
    assert c2 == cert_path
    assert k2 == key_path


def test_bridge_session_registration_and_lookup():
    bridge = WordBridge()
    session = bridge.register_session(
        doc_url="https://vingroupjsc.sharepoint.com/sites/VF_AIDV/Shared%20Documents/Specs.docx",
        doc_title="Specs.docx",
        platform="web",
        client_id="test-client-1",
    )
    assert session.session_id.startswith("w_")
    assert session.doc_title == "Specs.docx"
    assert bridge.get_session(session.session_id) is session

    # Lookup by exact title / filename
    found = bridge.find_session_for_document("Specs.docx")
    assert found is session

    # Lookup by URL
    found2 = bridge.find_session_for_document("https://vingroupjsc.sharepoint.com/sites/VF_AIDV/Shared%20Documents/Specs.docx")
    assert found2 is session

    # Lookup by partial URL
    found3 = bridge.find_session_for_document("Shared%20Documents/Specs.docx", item_name="Specs.docx")
    assert found3 is session

    # Heartbeat updates timestamp
    old_hb = session.last_heartbeat
    assert bridge.heartbeat_session(session.session_id, doc_title="Specs_v2.docx")
    assert session.doc_title == "Specs_v2.docx"
    assert session.last_heartbeat >= old_hb


def test_bridge_http_server_endpoints(tmp_path):
    cert_path, key_path = get_or_create_dev_certificate(cert_dir=tmp_path)
    bridge = WordBridge()
    # Use HTTP mode for simple socket testing without SSL verification hurdles in tests
    bridge.ensure_running(
        host="127.0.0.1",
        port=13650,
        ssl_enabled=False,
        cert_file=str(cert_path),
        key_file=str(key_path),
    )
    try:
        base = bridge.base_url
        assert "http://127.0.0.1:13650" == base

        # 1. GET /api/status
        with urllib.request.urlopen(f"{base}/api/status", timeout=5) as resp:
            assert resp.status == 200
            data = json.loads(resp.read().decode("utf-8"))
            assert data["status"] == "running"
            assert data["port"] == 13650

        # 2. GET /manifest.xml
        with urllib.request.urlopen(f"{base}/manifest.xml", timeout=5) as resp:
            assert resp.status == 200
            xml = resp.read().decode("utf-8")
            assert "127.0.0.1:13650" in xml
            assert "<OfficeApp" in xml

        # 3. GET /taskpane.html
        with urllib.request.urlopen(f"{base}/taskpane.html", timeout=5) as resp:
            assert resp.status == 200
            html = resp.read().decode("utf-8")
            assert "Auto 365 Companion" in html

        # 4. POST /api/sessions/register
        reg_req = urllib.request.Request(
            f"{base}/api/sessions/register",
            data=json.dumps({
                "doc_url": "https://example.sharepoint.com/TestDoc.docx",
                "doc_title": "TestDoc.docx",
                "client_id": "client-1",
            }).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(reg_req, timeout=5) as resp:
            assert resp.status == 200
            reg_data = json.loads(resp.read().decode("utf-8"))
            sess_id = reg_data["session_id"]
            token = reg_data["token"]
            assert sess_id
            assert token

        # 5. POST heartbeat
        hb_req = urllib.request.Request(
            f"{base}/api/sessions/{sess_id}/heartbeat",
            data=json.dumps({"doc_title": "TestDoc.docx"}).encode("utf-8"),
            headers={"Content-Type": "application/json", "X-Session-Token": token},
            method="POST",
        )
        with urllib.request.urlopen(hb_req, timeout=5) as resp:
            assert resp.status == 200

        # 6. Job submission and polling
        session = bridge.get_session(sess_id)
        assert session is not None

        # Start a thread to submit job
        job_results = []
        def run_job():
            res = bridge.execute_comments_live(
                session,
                comments=[{"anchor": "section 1", "text": "Needs review"}],
                author="TestReviewer",
                timeout=5.0,
            )
            job_results.extend(res)

        import threading
        job_thread = threading.Thread(target=run_job, daemon=True)
        job_thread.start()

        # Poll the job from the mock add-in
        with urllib.request.urlopen(f"{base}/api/sessions/{sess_id}/jobs/poll?timeout=3", timeout=5) as resp:
            assert resp.status == 200
            job_payload = json.loads(resp.read().decode("utf-8"))
            assert job_payload["action"] == "add_comments"
            assert len(job_payload["comments"]) == 1
            job_id = job_payload["job_id"]

        # Post job result from mock add-in
        res_req = urllib.request.Request(
            f"{base}/api/sessions/{sess_id}/jobs/{job_id}/result",
            data=json.dumps({
                "status": "ok",
                "results": [{
                    "anchor": "section 1",
                    "text": "Needs review",
                    "status": "ok",
                    "comment_id": "c1001",
                    "matched_text": "section 1",
                }],
            }).encode("utf-8"),
            headers={"Content-Type": "application/json", "X-Session-Token": token},
            method="POST",
        )
        with urllib.request.urlopen(res_req, timeout=5) as resp:
            assert resp.status == 200

        job_thread.join(timeout=3.0)
        assert len(job_results) == 1
        assert job_results[0]["comment_id"] == "c1001"
    finally:
        bridge.stop()

@pytest.mark.anyio
async def test_add_sharepoint_docx_comments_live_mode_refusal(monkeypatch):
    """When mode='live' is requested and no Word session is connected, fail immediately."""
    import tools as tools_mod

    class FakeSP:
        def resolve_file(self, _path):
            return "drive_1", {"id": "item_1", "name": "Document.docx", "eTag": "tag1"}

    monkeypatch.setattr(tools_mod, "sp", lambda: FakeSP())

    class EmptyBridge:
        is_running = True
        def find_session_for_document(self, *args, **kwargs):
            return None

    monkeypatch.setattr(tools_mod, "get_bridge", lambda: EmptyBridge())

    mcp = MCPServer("test")
    tools_mod.register_all(mcp)
    with pytest.raises(ToolError) as excinfo:
        await mcp.call_tool(
            "add_sharepoint_docx_comments",
            {
                "file_url_or_guid": "Document.docx",
                "comments": [{"anchor": "foo", "text": "bar"}],
                "mode": "live",
                "is_user_confirm": True,
            },
        )
    assert "Chưa có phiên Word Companion Add-in" in str(excinfo.value)


@pytest.mark.anyio
async def test_add_sharepoint_docx_comments_live_mode_success(monkeypatch):
    """When a Word session is connected, add_sharepoint_docx_comments uses live Word API."""
    import tools as tools_mod

    class FakeSP:
        def resolve_file(self, _path):
            return "drive_1", {"id": "item_1", "name": "Roadmap.docx", "eTag": "tag1", "webUrl": "https://example.com/Roadmap.docx"}
        def put_file_bytes(self, *args, **kwargs):
            raise AssertionError("put_file_bytes must NOT be called in live Word mode")

    monkeypatch.setattr(tools_mod, "sp", lambda: FakeSP())

    class MockLiveBridge:
        is_running = True
        def find_session_for_document(self, *args, **kwargs):
            return WordSession(
                session_id="w_test1",
                token="tok1",
                doc_url="https://example.com/Roadmap.docx",
                doc_title="Roadmap.docx",
                platform="web",
                client_id="c1",
            )
        def execute_comments_live(self, session, comments, author="", timeout=35.0):
            return [
                {
                    "anchor": c["anchor"],
                    "text": c["text"],
                    "status": "ok",
                    "comment_id": "cmt_99",
                    "matched_text": c["anchor"],
                }
                for c in comments
            ]

    monkeypatch.setattr(tools_mod, "get_bridge", lambda: MockLiveBridge())

    mcp = MCPServer("test")
    tools_mod.register_all(mcp)

    # 1. Call with is_user_confirm=False -> raises ApprovalRequiredError with live draft
    with pytest.raises(ToolError) as excinfo:
        await mcp.call_tool(
            "add_sharepoint_docx_comments",
            {
                "file_url_or_guid": "Roadmap.docx",
                "comments": [{"anchor": "Milestone 1", "text": "Review scope"}],
                "is_user_confirm": False,
            },
        )
    refusal = str(excinfo.value)
    assert "CHƯA GỬI" in refusal
    assert "Roadmap.docx" in refusal
    assert "Milestone 1" in refusal
    assert "Review scope" in refusal

    # 2. Call with is_user_confirm=True -> succeeds directly without put_file_bytes
    res = await mcp.call_tool(
        "add_sharepoint_docx_comments",
        {
            "file_url_or_guid": "Roadmap.docx",
            "comments": [{"anchor": "Milestone 1", "text": "Review scope"}],
            "is_user_confirm": True,
        },
    )
    result_text = str(res)
    assert "✓ Đã thêm 1 comment trực tiếp vào `Roadmap.docx`" in result_text
    assert "Word Companion Add-in" in result_text
    assert "Comment ID: `cmt_99`" in result_text


@pytest.mark.anyio
async def test_get_word_companion_status_tool(monkeypatch):
    import tools as tools_mod

    class MockBridge:
        is_running = True
        def get_status_summary(self):
            return {
                "status": "running",
                "base_url": "https://127.0.0.1:3650",
                "port": 3650,
                "ssl": True,
            }
        def get_active_sessions_info(self):
            return [{
                "session_id": "w_abc123",
                "doc_title": "Architecture.docx",
                "doc_url": "https://vingroup.sharepoint.com/Architecture.docx",
                "platform": "web",
                "last_seen_ago_s": 2,
            }]

    monkeypatch.setattr(tools_mod, "get_bridge", lambda: MockBridge())

    mcp = MCPServer("test")
    tools_mod.register_all(mcp)

    res = await mcp.call_tool("get_word_companion_status", {})
    output = str(res)
    assert "Word Companion Bridge" in output
    assert "Architecture.docx" in output
    assert "w_abc123" in output
