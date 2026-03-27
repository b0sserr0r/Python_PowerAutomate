import os
import sys
import time
import json
import uuid
import datetime as dt
from typing import Optional, List, Dict

import requests
from msal import ConfidentialClientApplication
from dotenv import load_dotenv

ISO_FMT = "%Y-%m-%dT%H:%M:%S.%fZ"
GUID_RE = r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"

STATUS_MAP = {
    0: "NotSpecified",
    1: "Paused",
    2: "Running",
    3: "Waiting",
    4: "Succeeded",
    5: "Skipped",
    6: "Suspended",
    7: "Cancelled",
    8: "Failed",
}


def utcnow() -> dt.datetime:
    return dt.datetime.now(dt.timezone.utc)


def parse_dt(s: str) -> dt.datetime:
    try:
        # Dataverse usually returns Zulu times like 2026-03-27T06:08:09Z or with millis
        if s.endswith("Z") and "." not in s:
            return dt.datetime.strptime(s, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=dt.timezone.utc)
        return dt.datetime.strptime(s, ISO_FMT).replace(tzinfo=dt.timezone.utc)
    except Exception:
        # Fallback: try fromisoformat (Python 3.11+ handles Z poorly), strip Z
        return dt.datetime.fromisoformat(s.replace("Z", "+00:00"))


def get_env(name: str, required: bool = True, default: Optional[str] = None) -> str:
    val = os.getenv(name, default)
    if required and not val:
        print(f"Missing required env var: {name}")
        sys.exit(2)
    return val


def acquire_token(tenant_id: str, client_id: str, client_secret: str, resource: str) -> str:
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = ConfidentialClientApplication(client_id=client_id, authority=authority, client_credential=client_secret)
    scope = [f"{resource}/.default"]

    result = app.acquire_token_silent(scopes=scope, account=None)
    if not result:
        result = app.acquire_token_for_client(scopes=scope)
    if "access_token" not in result:
        raise RuntimeError(f"Failed to acquire token: {result}")
    return result["access_token"]


def list_desktop_workflows(resource: str, token: str) -> List[Dict[str, str]]:
    url = f"{resource}/api/data/v9.2/workflows?$filter=category eq 6&$select=name,workflowid&$orderby=name"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
    }
    items: List[Dict[str, str]] = []
    next_link = url
    while next_link:
        resp = requests.get(next_link, headers=headers, timeout=60)
        if resp.status_code != 200:
            raise RuntimeError(f"Failed to list workflows: {resp.status_code} {resp.text}")
        data = resp.json()
        items.extend(data.get("value", []))
        next_link = data.get("@odata.nextLink")
    # Normalize fields
    result: List[Dict[str, str]] = []
    for it in items:
        name = str(it.get("name") or "").strip()
        wid = str(it.get("workflowid") or "").strip("{} ")
        if wid:
            result.append({"name": name or wid, "workflowid": wid})
    return result


def prompt_select_workflow(flows: List[Dict[str, str]]) -> Optional[str]:
    if not flows:
        print("No Desktop Workflows (category=6) found.")
        return None
    print("Available Desktop Workflows:")
    for idx, f in enumerate(flows, start=1):
        print(f"{idx:3d}. {f.get('name')}    [{f.get('workflowid')}]")
    while True:
        sel = input("Enter number to run (or 'q' to quit): ").strip()
        if sel.lower() in {"q", "quit", "exit"}:
            return None
        if not sel.isdigit():
            print("Please enter a valid number.")
            continue
        i = int(sel)
        if 1 <= i <= len(flows):
            return flows[i - 1]["workflowid"]
        print("Number out of range. Try again.")


def prompt_inputs_payload() -> Optional[dict]:
    print("\nOptional: provide additional JSON inputs for this run.")
    print("- Enter a file path to a .json, or paste raw JSON.")
    print("- Press Enter to skip.")
    while True:
        s = input("Inputs (path or JSON): ").strip()
        if not s:
            return None
        if os.path.exists(s):
            try:
                with open(s, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as ex:
                print(f"Failed to load file: {ex}")
                continue
        try:
            return json.loads(s)
        except Exception as ex:
            print(f"Invalid JSON: {ex}")
            yn = input("Try again? [y/N]: ").strip().lower()
            if yn not in ("y", "yes"):
                return None


def _deep_merge(a: dict, b: dict) -> dict:
    out = dict(a or {})
    for k, v in (b or {}).items():
        if k in out and isinstance(out[k], dict) and isinstance(v, dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def normalize_action_payload(p: Optional[dict]) -> dict:
    q = dict(p or {})
    # Handle 'inputs': remove if empty; stringify if object/array
    if "inputs" in q:
        v = q["inputs"]
        should_remove = (
            v is None or
            (isinstance(v, (dict, list)) and len(v) == 0) or
            (isinstance(v, str) and (v.strip() == "" or v.strip() in ("{}", "[]")))
        )
        if should_remove:
            q.pop("inputs", None)
        elif not isinstance(v, str):
            try:
                q["inputs"] = json.dumps(v, ensure_ascii=False)
            except Exception:
                # Leave as-is; API will return a clear error if invalid
                pass
    # Normalize runMode capitalization if provided
    rm = q.get("runMode")
    if isinstance(rm, str):
        rml = rm.strip().lower()
        if rml in ("attended", "unattended"):
            q["runMode"] = rml.capitalize()
    return q


def call_run_desktop_flow(resource: str, token: str, workflow_id: str, payload: Optional[dict] = None) -> requests.Response:
    url = f"{resource}/api/data/v9.2/workflows({workflow_id})/Microsoft.Dynamics.CRM.RunDesktopFlow"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json; charset=utf-8",
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        # Correlation to help find the run if needed
        "MS-Client-Request-Id": str(uuid.uuid4()),
    }
    data = json.dumps(payload or {})
    resp = requests.post(url, headers=headers, data=data, timeout=60)
    return resp


def try_extract_flowsession_id_from_response(resp: requests.Response) -> Optional[str]:
    try:
        body = resp.json()
        if isinstance(body, dict):
            for k, v in body.items():
                lk = str(k).lower()
                if lk in ("flowsessionid", "flow_session_id", "flowSessionId".lower()):
                    if isinstance(v, str):
                        return v.strip("{}")
                if isinstance(v, str) and len(v) >= 36 and "-" in v:
                    import re
                    m = re.search(GUID_RE, v)
                    if m:
                        return m.group(0)
        if isinstance(body, list) and body:
            cand = body[0]
            if isinstance(cand, dict):
                v = cand.get("flowsessionid") or cand.get("flowSessionId")
                if isinstance(v, str):
                    return v.strip("{}")
    except Exception:
        pass
    for h in ("OData-EntityId", "Location"):
        val = resp.headers.get(h)
        if not val:
            continue
        import re
        m = re.search(GUID_RE, val)
        if m:
            return m.group(0)
    return None


def get_flowsession(resource: str, token: str, flowsession_id: str) -> Optional[dict]:
    url = (
        f"{resource}/api/data/v9.2/flowsessions({flowsession_id})?"
        f"$select=statuscode,statecode,startedon,completedon,errorcode,errordetails"
    )
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
    }
    resp = requests.get(url, headers=headers, timeout=60)
    if resp.status_code == 200:
        return resp.json()
    return None


def monitor_flowsession(resource: str, token: str, flowsession_id: str, poll_interval: int, timeout_sec: int, verbose: bool = False) -> dict:
    deadline = time.time() + timeout_sec
    last_status_code = None
    while time.time() < deadline:
        fs = get_flowsession(resource, token, flowsession_id)
        if fs:
            scode = fs.get("statuscode")
            try:
                scode_int = int(scode)
            except Exception:
                scode_int = None
            status_name = STATUS_MAP.get(scode_int, str(scode))
            if scode_int != last_status_code:
                print(f"Status: {status_name}")
            last_status_code = scode_int
            if scode_int == 4:
                return {"outcome": "Succeeded", "flowsession": fs}
            if scode_int in (7, 8):
                return {"outcome": "Failed", "flowsession": fs}
        else:
            if verbose:
                print("FlowSession not found yet or not accessible.")
        time.sleep(poll_interval)
    return {"outcome": "TimedOut", "flowsession": fs if 'fs' in locals() else None}


def get_flowsession_outputs(resource: str, token: str, flowsession_id: str) -> Optional[object]:
    url = f"{resource}/api/data/v9.2/flowsessions({flowsession_id})/outputs/$value"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json, text/plain, */*",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
    }
    resp = requests.get(url, headers=headers, timeout=60)
    if resp.status_code != 200:
        return None
    text = resp.text.strip()
    if not text:
        return None
    try:
        return json.loads(text)
    except Exception:
        return text


def find_recent_flowsession_for_workflow(resource: str, token: str, workflow_id: str, started_after: dt.datetime, tolerance_sec: int) -> Optional[dict]:
    window_start = (started_after - dt.timedelta(seconds=tolerance_sec)).isoformat()
    url = (
        f"{resource}/api/data/v9.2/flowsessions?"
        f"$select=flowsessionid,statuscode,startedon,completedon&"
        f"$filter=workflowid eq '{workflow_id}' and startedon ge {window_start}&$orderby=startedon desc&$top=3"
    )
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
    }
    resp = requests.get(url, headers=headers, timeout=60)
    if resp.status_code != 200:
        return None
    value = resp.json().get("value", [])
    return value[0] if value else None


def find_latest_flowrun_for_workflow(resource: str, token: str, workflow_id: str) -> Optional[dict]:
    # Filter by string property 'workflowid' for reliability
    url = (
        f"{resource}/api/data/v9.2/flowruns?"
        f"$select=flowrunid,workflowid,status,errormessage,errorcode,createdon,starttime,endtime&"
        f"$filter=workflowid eq '{workflow_id}'&$orderby=createdon desc&$top=3"
    )
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
    }
    resp = requests.get(url, headers=headers, timeout=60)
    if resp.status_code != 200:
        raise RuntimeError(f"Failed to query flowruns: {resp.status_code} {resp.text}")
    value = resp.json().get("value", [])
    return value[0] if value else None


def monitor_run(resource: str, token: str, workflow_id: str, started_after: dt.datetime, poll_interval: int, timeout_sec: int, start_tolerance_sec: int, verbose: bool = False) -> dict:
    deadline = time.time() + timeout_sec
    last_status = None
    while time.time() < deadline:
        fr = find_latest_flowrun_for_workflow(resource, token, workflow_id)
        if fr:
            createdon = fr.get("createdon")
            if createdon:
                createdon_dt = parse_dt(createdon)
                if verbose:
                    print(f"Latest run createdon={createdon} status={fr.get('status')} id={fr.get('flowrunid')}")
                # Accept runs that started after our trigger minus tolerance for clock skew and queueing
                if createdon_dt >= started_after - dt.timedelta(seconds=start_tolerance_sec):
                    status = (fr.get("status") or "").lower()
                    if status != last_status:
                        print(f"Status: {fr.get('status')} (flowrunid={fr.get('flowrunid')})")
                        last_status = status
                    if status in {"succeeded", "success", "completed", "complete"}:
                        return {"outcome": "Succeeded", "flowrun": fr}
                    if status in {"failed", "error", "cancelled", "canceled"}:
                        return {"outcome": "Failed", "flowrun": fr}
            else:
                if verbose:
                    print("No 'createdon' on latest flowrun record.")
        else:
            if verbose:
                print("No flowrun found yet for this workflow.")
        time.sleep(poll_interval)
    return {"outcome": "TimedOut", "flowrun": fr if 'fr' in locals() else None}


def load_optional_inputs(path: Optional[str]) -> Optional[dict]:
    if not path:
        return None
    if not os.path.exists(path):
        raise FileNotFoundError(f"Inputs file not found: {path}")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def main():
    load_dotenv()

    tenant_id = get_env("TENANT_ID")
    client_id = get_env("CLIENT_ID")
    client_secret = get_env("CLIENT_SECRET")
    dataverse_url = get_env("DATAVERSE_URL")
    workflow_id = os.getenv("WORKFLOW_ID", "").strip()

    # Normalize URL; workflow id will be handled later (may be picked interactively)
    dataverse_url = dataverse_url.rstrip("/")

    poll_interval = int(get_env("POLL_INTERVAL_SEC", required=False, default="5"))
    timeout_sec = int(get_env("POLL_TIMEOUT_SEC", required=False, default="1200"))
    start_tolerance_sec = int(get_env("START_TOLERANCE_SEC", required=False, default="300"))
    verbose = os.getenv("VERBOSE", "0") == "1"

    # Optional CLI args
    inputs_path = None
    run_mode_cli = None
    connection_name_cli = None
    flowsession_id_cli = None
    pick_cli = False
    if len(sys.argv) > 1:
        args = sys.argv[1:]
        i = 0
        while i < len(args):
            if args[i] in ("-i", "--inputs") and i + 1 < len(args):
                inputs_path = args[i + 1]
                i += 2
            elif args[i] in ("-p", "--pick", "--select"):
                pick_cli = True
                i += 1
            elif args[i] in ("-m", "--run-mode") and i + 1 < len(args):
                run_mode_cli = args[i + 1]
                i += 2
            elif args[i] in ("-c", "--connection-name") and i + 1 < len(args):
                connection_name_cli = args[i + 1]
                i += 2
            elif args[i] in ("-s", "--flowsession-id") and i + 1 < len(args):
                flowsession_id_cli = args[i + 1]
                i += 2
            else:
                print("Usage: python run_desktop_flow.py [--pick] [--run-mode Attended|Unattended] [--connection-name <name>] [--flowsession-id <id>] [--inputs inputs.json]")
                sys.exit(1)

    try:
        token = acquire_token(tenant_id, client_id, client_secret, dataverse_url)
    except Exception as ex:
        print(f"Auth error: {ex}")
        sys.exit(3)

    payload = None
    if inputs_path:
        try:
            payload = load_optional_inputs(inputs_path)
        except Exception as ex:
            print(f"Failed to load inputs: {ex}")
            sys.exit(4)

    # Inject required parameters if missing
    run_mode_env = os.getenv("RUN_MODE")
    connection_name_env = os.getenv("CONNECTION_NAME")
    run_mode = (run_mode_cli or run_mode_env or "").strip()
    connection_name = (connection_name_cli or connection_name_env or "").strip()

    if payload is None:
        payload = {}
    if "runMode" not in payload and run_mode:
        payload["runMode"] = run_mode
    if "connectionName" not in payload and connection_name:
        payload["connectionName"] = connection_name
    # If still missing, that's okay — some flows may not require them.

    # If user asked to pick, or WORKFLOW_ID missing/placeholder, list and select
    interactive_pick = pick_cli or not workflow_id or workflow_id.upper() in {"SELECT", "PICK", "PROMPT"}
    if interactive_pick:
        try:
            flows = list_desktop_workflows(dataverse_url, token)
        except Exception as ex:
            print(f"Failed to list Desktop Workflows: {ex}")
            sys.exit(12)
        sel_id = prompt_select_workflow(flows)
        if not sel_id:
            print("No workflow selected. Exiting.")
            sys.exit(0)
        workflow_id = sel_id

        # After selecting a workflow, allow user to provide additional JSON inputs
        extra = prompt_inputs_payload()
        if extra:
            payload = _deep_merge(payload or {}, extra)

    # Normalize selected/entered GUID
    workflow_id = workflow_id.strip()
    if workflow_id.startswith("{") and workflow_id.endswith("}"):
        workflow_id = workflow_id[1:-1]

    # Validate GUID format (basic)
    try:
        uuid.UUID(workflow_id)
    except Exception:
        print("WORKFLOW_ID must be a valid GUID, e.g. 1015b2f8-5575-45dd-b1ba-adca4f1f5957")
        sys.exit(2)

    flowsession_id = (flowsession_id_cli or os.getenv("FLOWSESSION_ID") or "").strip("{} ")
    started_at = utcnow()

    # Proceed even if runMode/connectionName not provided; API may accept empty or inputs-only payloads.

    if not flowsession_id:
        print("Triggering Desktop Flow run...")
        try:
            resp = call_run_desktop_flow(dataverse_url, token, workflow_id, normalize_action_payload(payload))
        except requests.RequestException as ex:
            print(f"HTTP error calling RunDesktopFlow: {ex}")
            sys.exit(5)

        if resp.status_code not in (200, 202, 204):
            print(f"RunDesktopFlow failed: {resp.status_code}\n{resp.text}")
            sys.exit(6)

        flowsession_id = try_extract_flowsession_id_from_response(resp) or flowsession_id
        print(f"RunDesktopFlow accepted (HTTP {resp.status_code}).")

    if not flowsession_id:
        recent = find_recent_flowsession_for_workflow(dataverse_url, token, workflow_id, started_at, start_tolerance_sec)
        if recent and recent.get("flowsessionid"):
            flowsession_id = str(recent.get("flowsessionid")).strip("{} ")
            if verbose:
                print(f"Discovered recent FlowSession: {flowsession_id}")

    if flowsession_id:
        print(f"Monitoring FlowSession {flowsession_id} until completion...")
        result = monitor_flowsession(dataverse_url, token, flowsession_id, poll_interval, timeout_sec, verbose)
    else:
        print("FlowSession Id not found. Falling back to flowrun-based monitoring...")
        result = monitor_run(dataverse_url, token, workflow_id, started_at, poll_interval, timeout_sec, start_tolerance_sec, verbose)

    outcome = result["outcome"]
    fr = result.get("flowrun") or {}
    fs = result.get("flowsession") or {}

    print("\n=== Final Result ===")
    print(f"Outcome: {outcome}")
    if fr:
        print(f"FlowRunId: {fr.get('flowrunid')}")
        print(f"Status: {fr.get('status')}")
        if fr.get("errorcode") or fr.get("errormessage"):
            print(f"ErrorCode: {fr.get('errorcode')}")
            print(f"ErrorMessage: {fr.get('errormessage')}")
        if fr.get("starttime"):
            print(f"Start: {fr.get('starttime')}")
        if fr.get("endtime"):
            print(f"End: {fr.get('endtime')}")
    if fs:
        print(f"FlowSessionId: {flowsession_id}")
        sc = fs.get('statuscode')
        try:
            sc_int = int(sc)
        except Exception:
            sc_int = None
        print(f"Status: {STATUS_MAP.get(sc_int, sc) if sc_int is not None else sc}")
        if fs.get('startedon'):
            print(f"StartedOn: {fs.get('startedon')}")
        if fs.get('completedon'):
            print(f"CompletedOn: {fs.get('completedon')}")
        if fs.get('errorcode') or fs.get('errordetails'):
            print(f"ErrorCode: {fs.get('errorcode')}")
            print(f"ErrorDetails: {fs.get('errordetails')}")
        # Try to fetch outputs only when DISPLAY_OUTPUT=true
        show_outputs = os.getenv("DISPLAY_OUTPUT", "").strip().lower() in ("1", "true", "yes", "y")
        if show_outputs:
            outs = get_flowsession_outputs(dataverse_url, token, flowsession_id)
            if outs is not None:
                print("Outputs:")
                if isinstance(outs, dict):
                    for k, v in outs.items():
                        print(f"- {k}: {v}")
                else:
                    print(outs)

    # Exit code conventions
    if outcome == "Succeeded":
        sys.exit(0)
    elif outcome == "Failed":
        sys.exit(10)
    else:
        sys.exit(11)


if __name__ == "__main__":
    main()
