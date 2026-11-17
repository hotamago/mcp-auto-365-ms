"""Native forward payloads, using anonymised shapes read from Teams UI messages.

Real samples: single/file 1793342958583, note 1793352099595, multi 1790672530783.
All client and tool tests are offline; no browser session is accessed.
"""

from __future__ import annotations

import copy
import json
import urllib.parse

import pytest
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

import tools as tools_mod
from common.errors import ConfigError, Mcp365Error
from teams.client import TeamsClient, build_forward_message, parse_attachments, parse_quoted, text_to_teams_html

SOURCE = "19:source@thread.v2"
TARGET = "48:notes"
SENDER = "8:orgid:11111111-1111-1111-1111-111111111111"
FILE = {
    "fileInfo": {
        "serverRelativeUrl": "/personal/test/Documents/Microsoft Teams Chat Files/test.md",
        "shareID": "share-id",
        "fileUrl": "https://tenant-my.sharepoint.com/personal/test/Documents/test.md",
        "shareUrl": "https://tenant-my.sharepoint.com/:t:/g/personal/test/share-id",
        "siteUrl": "https://tenant-my.sharepoint.com/personal/test",
    },
    "fileName": "test.md", "fileType": "md", "itemid": "file-id",
}


def original(msg_id="1793262588641", content="<p>Full <b>original</b> &amp; text.</p>", **properties):
    return {
        "id": msg_id, "messagetype": "RichText/Html", "content": content,
        "from": f"https://teams.example/v1/users/ME/contacts/{SENDER}",
        "imdisplayname": "Original Sender",
        "originalarrivaltime": "2026-10-29T08:29:48.6410000Z",
        "properties": properties,
    }


def test_single_matches_observed_native_html_and_properties():
    raw = original(files=json.dumps([FILE]), mentions='[{"mri":"do-not-notify"}]', importance="high")
    before = copy.deepcopy(raw)
    content, props = build_forward_message(SOURCE, [raw])
    assert content == '<blockquote itemtype="http://schema.skype.com/Forward">' + raw["content"] + "</blockquote>"
    assert props == {
        "mentions": "[]", "forwardTemplateId": "basic_forward_message_template",
        "originalMessageContext": {
            "originalSender": SENDER, "originalSentTime": "2026-10-29T08:29:48.641Z",
            "messageId": 1793262588641, "originalThreadId": SOURCE,
            "threadType": "chat", "messageType": "RichText/Html",
        },
        "files": json.dumps([FILE], ensure_ascii=False),
    }
    assert raw == before
    parsed = parse_quoted({"content": content, "properties": props})
    assert [(p["kind"], p["message_id"], p["conversation_id"], p["sender_mri"]) for p in parsed] == [
        ("forward", raw["id"], SOURCE, SENDER)
    ]
    assert parse_attachments({"properties": props})[0]["url"] == FILE["fileInfo"]["fileUrl"]


def test_multi_keeps_input_order_full_html_files_and_note_converter():
    note = "**For review**: https://example.test/a. <script>"
    raw = original(content="<div>First<br>line</div>", files=[FILE])
    second = original("1793262588640", "<p>Second</p>", files=json.dumps([{**FILE, "fileName": "other.md"}]))
    content, props = build_forward_message(SOURCE, [raw, second], note)
    assert content.startswith(text_to_teams_html(note) + '<blockquote itemtype="http://schema.skype.com/Forward">')
    assert content.index("First") < content.index("Second")
    assert props["originalMessageContext"]["messageId"] == int(raw["id"])
    assert props["originalMessageContext1"]["messageId"] == int(second["id"])
    assert [f["fileName"] for f in json.loads(props["files"])] == ["test.md", "other.md"]
    assert props["additionalMessageContext"] == {"amsreferences": [], "cards": "[]", "links": "[]"}
    assert len(parse_quoted({"content": content, "properties": props})) == 2


def test_file_only_and_plain_text_are_not_lost_or_interpreted_as_markdown():
    content, props = build_forward_message(SOURCE, [original(content="", files=[FILE])])
    assert content.endswith('Forward"></blockquote>')
    assert len(parse_attachments({"properties": props})) == 1
    raw = original(content="a < b & **literal**\nnext")
    raw["messagetype"] = "Text"
    content, props = build_forward_message(SOURCE, [raw])
    assert "<p>a &lt; b &amp; **literal**<br/>next</p>" in content
    assert props["originalMessageContext"]["messageType"] == "Text"


def test_nested_reply_author_and_preview_are_preserved():
    body = (
        '<blockquote itemscope itemtype="http://schema.skype.com/Reply" itemid="1793262500000">'
        f'<strong itemprop="mri" itemid="{SENDER}">Original Sender</strong>'
        '<span itemprop="time" itemid="1793262500000"></span><p itemprop="preview">preview</p>'
        '</blockquote><p>Reply text</p>'
    )
    content, props = build_forward_message(SOURCE, [original(content=body)])
    assert body in content
    quoted = parse_quoted({"content": content, "properties": props})
    assert [q["kind"] for q in quoted] == ["forward", "reply"]
    assert quoted[1]["sender_mri"] == SENDER
    assert quoted[1]["conversation_id"] == SOURCE
    assert quoted[1]["preview"] == "preview"


def test_forward_of_forward_keeps_contexts_aligned_with_nested_blocks():
    inner_html, inner_props = build_forward_message("19:older@thread.v2", [original()])
    wrapped = original("1793262588999", inner_html, **inner_props)
    html, props = build_forward_message(SOURCE, [wrapped, original("1793262588000")])
    quoted = parse_quoted({"content": html, "properties": props})
    assert [(q["message_id"], q["conversation_id"]) for q in quoted] == [
        (wrapped["id"], SOURCE), ("1793262588641", "19:older@thread.v2"), ("1793262588000", SOURCE),
    ]


def test_images_cards_and_links_remain_available():
    raw = original(amsreferences=["ams-id"], cards='[{"id":"card"}]', links=[{"url": "https://example.test"}])
    _, props = build_forward_message(SOURCE, [raw])
    assert props["amsreferences"] == ["ams-id"]
    assert json.loads(props["cards"]) == [{"id": "card"}]
    assert json.loads(props["links"]) == [{"url": "https://example.test"}]


@pytest.mark.parametrize("change", [
    {"properties": {"deletetime": 1}}, {"messagetype": "ThreadActivity/AddMember"}, {"from": ""}, {"id": ""},
    {"properties": {"files": "bad json"}}, {"properties": {"files": {"bad": "shape"}}},
])
def test_unforwardable_original_is_refused(change):
    with pytest.raises(ConfigError):
        build_forward_message(SOURCE, [{**original(), **change}])


def test_empty_selection_and_timestamp_fallback():
    with pytest.raises(ConfigError):
        build_forward_message(SOURCE, [])
    raw = original()
    raw.pop("originalarrivaltime")
    _, props = build_forward_message(SOURCE, [raw])
    assert props["originalMessageContext"]["originalSentTime"] == "2026-10-29T08:29:48.641Z"


@pytest.fixture
def client(monkeypatch, identity):
    c = TeamsClient()
    monkeypatch.setattr(type(c), "identity", property(lambda self: identity))
    monkeypatch.setattr(c, "find_conversation", lambda val: {"id": val, "name": f"Chat {val}"})
    monkeypatch.setattr(c, "_get_raw_message", lambda _chat, msg_id: original(msg_id, files=[FILE]))
    return c


@pytest.mark.parametrize("ids, expected", [
    ("1793262588641", ["1793262588641"]),
    (["1793262588642", "1793262588641", "1793262588642"], ["1793262588642", "1793262588641"]),
    ('["1793262588642", "1793262588641"]', ["1793262588642", "1793262588641"]),
])
def test_prepare_reads_only_and_previews_all_originals(client, ids, expected):
    plan = client.prepare_forward(SOURCE, ids, TARGET, "**Note**")
    assert plan["message_ids"] == expected
    for msg_id in expected:
        assert msg_id in plan["preview"]
    for text in ("**Note**", "Full", "Original Sender", "test.md", FILE["fileInfo"]["shareUrl"]):
        assert text in plan["preview"]
    assert plan["target"]["id"] == TARGET


def test_prepare_fails_if_any_source_is_unreadable_or_wrong(client, monkeypatch):
    def read(_chat, msg_id):
        if msg_id == "1793262588642":
            raise Mcp365Error("403", "Open the source chat.")
        return original(msg_id)

    monkeypatch.setattr(client, "_get_raw_message", read)
    with pytest.raises(Mcp365Error, match="403"):
        client.prepare_forward(SOURCE, ["1793262588641", "1793262588642"], TARGET)
    monkeypatch.setattr(client, "_get_raw_message", lambda *_: original("1793262588000"))
    with pytest.raises(ConfigError, match="1793262588641"):
        client.prepare_forward(SOURCE, "1793262588641", TARGET)


def test_send_posts_prepared_content_and_recovers_message_id(client, monkeypatch):
    plan = client.prepare_forward(SOURCE, "1793262588641", TARGET, "Note")
    calls = []

    def post(method, path, **kwargs):
        calls.append((method, path, json.loads(kwargs["data"])))
        return {"OriginalArrivalTime": 1794000000000}

    monkeypatch.setattr(client, "_chat_json", post)
    monkeypatch.setattr(client, "_resolve_sent_id", lambda conv, client_id, data: str(data["OriginalArrivalTime"]))
    result = client.forward_message(plan)
    method, path, payload = calls[0]
    assert len(calls) == 1 and method == "POST"
    assert path == f"/users/ME/conversations/{urllib.parse.quote(TARGET)}/messages"
    assert payload["content"] == plan["content"]
    assert payload["properties"] == plan["properties"]
    assert payload["messagetype"] == "RichText/Html" and payload["contenttype"] == "text"
    assert payload["imdisplayname"] == client.identity.display_name
    assert result["message_id"] == "1794000000000"
    assert result["source_chat_id"] == SOURCE


@pytest.mark.parametrize("status", [403, 500, None])
def test_failed_send_is_propagated_without_a_second_post(client, monkeypatch, status):
    plan = client.prepare_forward(SOURCE, "1793262588641", TARGET)
    calls = []
    error = Mcp365Error("send failed", "Read the target before retrying.")
    error.http_status = status

    def fail(*args, **kwargs):
        calls.append(args)
        raise error

    monkeypatch.setattr(client, "_chat_json", fail)
    with pytest.raises(Mcp365Error) as exc:
        client.forward_message(plan)
    assert exc.value is error
    assert len(calls) == 1


def test_unprovisioned_direct_chat_uses_same_payload_after_provisioning(client, monkeypatch):
    direct = "19:11111111-1111-1111-1111-111111111111_22222222-2222-2222-2222-222222222222@unq.gbl.spaces"
    plan = client.prepare_forward(SOURCE, "1793262588641", direct)
    calls = []

    def post(method, path, **kwargs):
        calls.append((path, kwargs["data"]))
        if len(calls) == 1:
            error = Mcp365Error("404 LocationLookupFailed", "Provision the chat.")
            error.http_status = 404
            raise error
        return {"OriginalArrivalTime": 1794000000000}

    provisioned = []
    monkeypatch.setattr(client, "create_or_get_direct_chat", lambda val: provisioned.append(val) or val)
    monkeypatch.setattr(client, "_chat_json", post)
    monkeypatch.setattr(client, "_resolve_sent_id", lambda *_: "1794000000000")
    client.forward_message(plan)
    assert calls[0] == calls[1]
    assert provisioned == [direct]


@pytest.mark.anyio
async def test_tool_requires_approval_of_exact_plan_and_sends_only_when_confirmed(client, monkeypatch):
    monkeypatch.setattr(tools_mod, "teams", lambda: client)
    sent = []

    def send(plan):
        sent.append(plan)
        return {
            "source_chat_id": SOURCE, "forwarded_message_ids": plan["message_ids"], "message_id": "1794000000000",
            "conversation_id": TARGET, "conversation_name": "Notes",
        }

    monkeypatch.setattr(client, "forward_message", send)
    server = MCPServer("test-forward")
    tools_mod.register_teams_tools(server)
    args = {
        "source_chat": SOURCE, "message_ids": ["1793262588642", "1793262588641"],
        "target_chat_name_or_id": TARGET, "note": "Exact **note**", "is_user_confirm": False,
    }
    with pytest.raises(ToolError) as exc:
        await server.call_tool("forward_teams_message", args)
    draft = str(exc.value)
    for text in (SOURCE, TARGET, "1793262588642", "1793262588641", "Exact **note**", "Original Sender", "test.md"):
        assert text in draft
    assert not sent
    await server.call_tool("forward_teams_message", {**args, "is_user_confirm": True})
    assert len(sent) == 1
    assert sent[0]["message_ids"] == args["message_ids"]
    assert sent[0]["note"] == args["note"]
