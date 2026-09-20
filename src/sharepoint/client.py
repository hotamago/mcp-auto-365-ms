"""SharePoint & OneDrive Client for exploring and downloading raw binary files."""

import subprocess
import json
import re
import time
import urllib.parse
import urllib.request
import urllib.error
import os
from pathlib import Path
from typing import Dict, Any, Optional, List
from common.chrome_cookies import ChromeCookieDecryptor


class SharePointClient:
    def __init__(self):
        self._token: Optional[str] = None
        self._token_expires: float = 0.0
        self._site_cache: Dict[str, str] = {}
        self._drive_cache: Dict[str, str] = {}

    def get_token(self) -> str:
        """Get Microsoft Graph access token via az CLI with caching."""
        now = time.time()
        if self._token and now < (self._token_expires - 60):
            return self._token

        try:
            cmd = ['az', 'account', 'get-access-token', '--resource', 'https://graph.microsoft.com', '-o', 'json']
            res = subprocess.run(cmd, capture_output=True, text=True, check=True)
            data = json.loads(res.stdout)
            self._token = data['accessToken']
            exp = data.get('expires_on')
            if isinstance(exp, (int, float)):
                self._token_expires = float(exp)
            else:
                self._token_expires = now + 3000
            return self._token
        except Exception as e:
            raise RuntimeError(f"Failed to obtain Microsoft Graph token via az CLI: {e}")

    def call_graph(self, path: str, method: str = 'GET', body: Optional[dict] = None) -> dict:
        """Execute a request against Microsoft Graph API."""
        token = self.get_token()
        url = f"https://graph.microsoft.com/v1.0{path}" if not path.startswith("http") else path
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json"
        }
        data = None
        if body is not None:
            data = json.dumps(body).encode('utf-8')
            headers["Content-Type"] = "application/json"

        req = urllib.request.Request(url, data=data, headers=headers, method=method)
        try:
            with urllib.request.urlopen(req) as resp:
                raw = resp.read().decode('utf-8')
                return json.loads(raw) if raw else {}
        except urllib.error.HTTPError as e:
            err_body = e.read().decode('utf-8', errors='ignore')
            raise RuntimeError(f"Graph API error ({e.code}) on {url}: {err_body}")

    def parse_sharepoint_url(self, url: str) -> Dict[str, Any]:
        """Extract hostname, site_name, folder_path, and sourcedoc from a SharePoint URL."""
        parsed = urllib.parse.urlparse(url)
        hostname = parsed.netloc
        path = urllib.parse.unquote(parsed.path)
        query = urllib.parse.parse_qs(parsed.query)

        res = {
            "hostname": hostname,
            "site_name": None,
            "site_path": None,
            "folder_path": None,
            "sourcedoc": None,
            "file_name": None,
            "is_personal": "-my.sharepoint.com" in hostname,
            "type": "unknown"
        }

        if "sourcedoc" in query:
            sdoc = re.sub(r'[{}]', '', query["sourcedoc"][0])
            res["sourcedoc"] = sdoc
            res["type"] = "document"

        if "file" in query:
            res["file_name"] = query["file"][0]

        site_match = re.search(r'/sites/([^/]+)', path)
        if site_match:
            res["site_name"] = site_match.group(1)
            res["site_path"] = f"/sites/{site_match.group(1)}"

        if "id" in query:
            res["folder_path"] = query["id"][0]
            if not res["sourcedoc"]:
                res["type"] = "folder"

        if any(marker in path for marker in [":w:", ":x:", ":p:", ":b:"]) or path.endswith(('.docx', '.xlsx', '.pptx', '.pdf')):
            if not res["type"] or res["type"] == "unknown":
                res["type"] = "document"

        return res

    def get_site_id(self, hostname: str, site_path: str) -> str:
        key = f"{hostname}:{site_path}"
        if key in self._site_cache:
            return self._site_cache[key]

        data = self.call_graph(f"/sites/{hostname}:{site_path}")
        site_id = data["id"]
        self._site_cache[key] = site_id
        return site_id

    def get_default_drive_id(self, site_id: str) -> str:
        if site_id in self._drive_cache:
            return self._drive_cache[site_id]

        data = self.call_graph(f"/sites/{site_id}/drives")
        drives = data.get("value", [])
        if not drives:
            raise RuntimeError(f"No document libraries found for site {site_id}")

        drive_id = drives[0]["id"]
        self._drive_cache[site_id] = drive_id
        return drive_id

    def list_folder_contents(self, drive_id: str, relative_path: str) -> List[Dict[str, Any]]:
        clean_path = re.sub(r'^/sites/[^/]+/Shared Documents/?', '', relative_path)
        clean_path = clean_path.strip('/')

        encoded_path = urllib.parse.quote(clean_path)
        if clean_path:
            endpoint = f"/drives/{drive_id}/root:/{encoded_path}:/children"
        else:
            endpoint = f"/drives/{drive_id}/root/children"

        data = self.call_graph(endpoint)
        return data.get("value", [])

    def get_item_by_guid(self, drive_id: str, guid: str) -> Dict[str, Any]:
        return self.call_graph(f"/drives/{drive_id}/items/{guid}")

    def get_item_versions(self, drive_id: str, item_id_or_guid: str) -> List[Dict[str, Any]]:
        data = self.call_graph(f"/drives/{drive_id}/items/{item_id_or_guid}/versions")
        return data.get("value", [])

    def read_link(self, url: str, max_depth: int = 2) -> str:
        """Explore SharePoint link: folder structure or document metadata (no markdown conversion)."""
        info = self.parse_sharepoint_url(url)
        if not info["hostname"] or not info["site_name"]:
            return f"Unable to parse SharePoint URL: {url}"

        site_id = self.get_site_id(info["hostname"], info["site_path"])
        drive_id = self.get_default_drive_id(site_id)

        if info["type"] == "document" and info["sourcedoc"]:
            return self.describe_document(drive_id, info["sourcedoc"], info.get("file_name"))

        folder_path = info.get("folder_path") or ""
        return self._render_folder_tree(drive_id, folder_path, max_depth=max_depth)

    def describe_document(self, drive_id: str, guid: str, file_name: Optional[str] = None) -> str:
        """Return clean document metadata and version history."""
        try:
            item = self.get_item_by_guid(drive_id, guid)
        except Exception as e:
            return f"Error retrieving document metadata ({guid}): {e}"

        name = item.get("name", file_name or "Unknown")
        size = item.get("size", 0)
        created_by = item.get("createdBy", {}).get("user", {}).get("displayName", "Unknown")
        created_time = item.get("createdDateTime", "")
        modified_by = item.get("lastModifiedBy", {}).get("user", {}).get("displayName", "Unknown")
        modified_time = item.get("lastModifiedDateTime", "")
        parent_path = item.get("parentReference", {}).get("path", "")
        if "root:" in parent_path:
            parent_path = parent_path.split("root:")[-1]

        out = [
            f"# Document: {name}",
            f"- **Path in SharePoint**: `{parent_path}/{name}`",
            f"- **GUID / UniqueId**: `{guid}`",
            f"- **Size**: {size:,} bytes ({size / (1024*1024):.2f} MB)",
            f"- **Created By**: {created_by} ({created_time})",
            f"- **Last Modified By**: {modified_by} ({modified_time})",
            f"- **Direct Web URL**: {item.get('webUrl')}",
            "",
            "## Version History"
        ]

        try:
            versions = self.get_item_versions(drive_id, guid)
            if versions:
                out.append("| Version | Modified Time | Modified By | Size |")
                out.append("| --- | --- | --- | --- |")
                for v in versions[:20]:
                    v_id = v.get("id")
                    v_mod = v.get("lastModifiedDateTime", "")[:19].replace("T", " ")
                    v_user = v.get("lastModifiedBy", {}).get("user", {}).get("displayName", "Unknown")
                    v_size = f"{v.get('size', 0):,} B"
                    out.append(f"| {v_id} | {v_mod} | {v_user} | {v_size} |")
                if len(versions) > 20:
                    out.append(f"| ... | and {len(versions) - 20} older versions | | |")
            else:
                out.append("*No version history available*")
        except Exception:
            out.append("*Unable to retrieve version history*")

        out.append(f"\n> **Download**: Call `download_sharepoint_link('{guid}')` to download the original file.")
        return "\n".join(out)

    def _render_folder_tree(self, drive_id: str, folder_path: str, max_depth: int = 2) -> str:
        out = [f"# SharePoint Folder: `{folder_path or '/'}`\n"]

        def _traverse(rel_path: str, depth: int, prefix: str = ""):
            if depth > max_depth:
                return
            try:
                items = self.list_folder_contents(drive_id, rel_path)
            except Exception as e:
                out.append(f"{prefix}- *[Error loading: {e}]*")
                return

            folders = [it for it in items if "folder" in it]
            files = [it for it in items if "folder" not in it]

            folders.sort(key=lambda x: x["name"].lower())
            files.sort(key=lambda x: x["name"].lower())

            for f in folders:
                cnt = f.get("folder", {}).get("childCount", 0)
                out.append(f"{prefix}📁 **{f['name']}/** *({cnt} items)*")
                sub_path = f"{rel_path}/{f['name']}".strip('/')
                _traverse(sub_path, depth + 1, prefix + "  ")

            for doc in files:
                sz = doc.get("size", 0)
                if sz > 1024 * 1024:
                    sz_str = f"{sz / (1024 * 1024):.1f} MB"
                elif sz > 1024:
                    sz_str = f"{sz / 1024:.1f} KB"
                else:
                    sz_str = f"{sz} B"
                mod = doc.get("lastModifiedDateTime", "")[:10]
                user = doc.get("lastModifiedBy", {}).get("user", {}).get("displayName", "")
                out.append(f"{prefix}📄 `{doc['name']}` *({sz_str}, {mod}{f', by {user}' if user else ''})*")

        _traverse(folder_path, 1)
        out.append("\n> **Download All**: Call `download_sharepoint_link(url)` to download all documents in this folder.")
        return "\n".join(out)

    def download_link(self, url_or_guid: str, target_dir: str = "docs/sharepoint") -> str:
        """Download a SharePoint file or an entire folder into target_dir."""
        parsed = urllib.parse.urlparse(url_or_guid) if url_or_guid.startswith("http") else None
        domain = parsed.netloc if parsed else "vingroupjsc.sharepoint.com"

        cookies = ChromeCookieDecryptor.get_cookies_for_domain(domain_pattern="sharepoint.com", cookie_names=['rtFa', 'FedAuth'])
        if not cookies.get('FedAuth') or not cookies.get('rtFa'):
            return f"Error: Could not retrieve SharePoint session cookies (rtFa/FedAuth) for domain {domain}."

        cookie_hdr = f"rtFa={cookies.get('rtFa')}; FedAuth={cookies.get('FedAuth')}"
        headers = {
            'Cookie': cookie_hdr,
            'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36',
            'Accept': '*/*'
        }

        guid_match = re.search(r'([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})', url_or_guid)
        target_path = Path(target_dir)
        target_path.mkdir(parents=True, exist_ok=True)

        downloaded = []

        def _download_direct(file_url: str, local_dst: Path) -> bool:
            local_dst.parent.mkdir(parents=True, exist_ok=True)
            req = urllib.request.Request(file_url, headers=headers)
            try:
                with urllib.request.urlopen(req) as resp:
                    data = resp.read()
                    local_dst.write_bytes(data)
                    downloaded.append((local_dst.name, len(data), str(local_dst)))
                    return True
            except Exception as e:
                downloaded.append((local_dst.name, 0, f"Error: {e}"))
                return False

        # 1. Personal OneDrive sharing links (-my.sharepoint.com/:x:/g/... or :w:/g/...)
        if url_or_guid.startswith("http") and "-my.sharepoint.com" in url_or_guid and any(marker in url_or_guid for marker in [":x:/", ":w:/", ":p:/", ":b:/"]):
            req = urllib.request.Request(url_or_guid, headers=headers)
            try:
                with urllib.request.urlopen(req) as resp:
                    html = resp.read().decode('utf-8', errors='ignore')
                m_wopi = re.search(r'var _wopiContextJson\s*=\s*({.*?});', html)
                if m_wopi:
                    wopi = json.loads(m_wopi.group(1))
                    file_name = wopi.get('FileName') or 'downloaded_file'
                    file_get_url = wopi.get('FileGetUrl')
                    if file_get_url:
                        _download_direct(file_get_url, target_path / file_name)
                    else:
                        downloaded.append((file_name, 0, "Error: FileGetUrl not found in WOPI context"))
                else:
                    downloaded.append(('shared_file', 0, "Error: _wopiContextJson not found in sharing page"))
            except Exception as e:
                downloaded.append(('shared_file', 0, f"Error fetching sharing link: {e}"))

        # 2. General SharePoint URL
        elif url_or_guid.startswith("http"):
            info = self.parse_sharepoint_url(url_or_guid)
            site_id = self.get_site_id(info["hostname"], info["site_path"])
            drive_id = self.get_default_drive_id(site_id)

            if info["type"] == "document" and info["sourcedoc"]:
                item = self.get_item_by_guid(drive_id, info["sourcedoc"])
                parent_path = item.get("parentReference", {}).get("path", "").split("root:")[-1]
                server_rel = f"/sites/{info['site_name']}/Shared Documents{parent_path}/{item['name']}"
                encoded = urllib.parse.quote(server_rel)
                file_url = f"https://{info['hostname']}{encoded}"
                local_file = target_path / item['name']
                _download_direct(file_url, local_file)
            else:
                folder_id = info.get("folder_path") or ""
                clean_path = re.sub(r'^/sites/[^/]+/Shared Documents/?', '', folder_id).strip('/')
                endpoint = f"/drives/{drive_id}/root:/{urllib.parse.quote(clean_path)}" if clean_path else f"/drives/{drive_id}/root"
                folder_item = self.call_graph(endpoint)

                def _sync(f_id: str, rel_sub: Path, s_parent: str):
                    items = self.call_graph(f"/drives/{drive_id}/items/{f_id}/children").get("value", [])
                    for it in items:
                        name = it['name']
                        if "folder" in it:
                            _sync(it['id'], rel_sub / name, f"{s_parent}/{name}")
                        else:
                            sz = it.get("size", 0)
                            if name.endswith('.mp4') and sz > 50*1024*1024:
                                continue
                            server_rel = f"{s_parent}/{name}"
                            encoded = urllib.parse.quote(server_rel)
                            file_url = f"https://{info['hostname']}{encoded}"
                            _download_direct(file_url, rel_sub / name)

                _sync(folder_item['id'], target_path, f"/sites/{info['site_name']}/Shared Documents/{clean_path}")

        elif guid_match:
            guid = guid_match.group(1)
            site_id = self.get_site_id("vingroupjsc.sharepoint.com", "/sites/VF_AIDV")
            drive_id = self.get_default_drive_id(site_id)
            item = self.get_item_by_guid(drive_id, guid)
            parent_path = item.get("parentReference", {}).get("path", "").split("root:")[-1]
            server_rel = f"/sites/VF_AIDV/Shared Documents{parent_path}/{item['name']}"
            encoded = urllib.parse.quote(server_rel)
            file_url = f"https://vingroupjsc.sharepoint.com{encoded}"
            local_file = target_path / item['name']
            _download_direct(file_url, local_file)
        else:
            return f"Invalid SharePoint link or GUID: {url_or_guid}"

        out = [f"# Downloaded {len([d for d in downloaded if d[1] > 0])} files to `{target_dir}`:\n"]
        out.append("| File Name | Size | Local Path |")
        out.append("| --- | --- | --- |")
        for name, sz, path in downloaded:
            if sz > 0:
                sz_str = f"{sz / (1024*1024):.2f} MB" if sz > 1024*1024 else f"{sz / 1024:.1f} KB"
                out.append(f"| `{name}` | {sz_str} | `{path}` |")
            else:
                out.append(f"| `{name}` | Failed | {path} |")

        return "\n".join(out)
