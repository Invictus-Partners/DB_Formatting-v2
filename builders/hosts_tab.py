# builders/hosts_tab.py
from pathlib import Path
from typing  import Dict

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import json

# ── 1. symbol ↔ priority maps (same as other tabs) ──────────────────────────
SYM2PRIO: Dict[str, int] = {"+": 4, "~": 3, "#": 2, "^": 1, "": 0, None: 0}
PRIO2SYM: Dict[int, str] = {v: k for k, v in SYM2PRIO.items()}

# ── 2. pure transform ───────────────────────────────────────────────────────
def build_hosts_df(
        df_virtual_devices: pd.DataFrame,
        physical_hosts_json: Path | str
) -> pd.DataFrame:
    """Return one‑row‑per‑host DataFrame with option roll‑ups & host specs."""

    # 2‑A  pull in host hardware / topology fields
    try:
        df_host_specs = pd.read_json(physical_hosts_json)
    except (FileNotFoundError, ValueError):
        # Physical Hosts.json missing or invalid: return empty hosts table
        # Define base headers for hosts
        desired_ids = [
            "cluster_name", "physical_device", "number_of_vms",
            "operating_system_type", "esx_version", "manufacturer",
            "model", "cpu_model", "total_number_of_processors",
            "cores_per_cpu", "total_number_of_cores", "oracle_core_factor",
        ]
        # Option columns from VM table (use df_virtual_devices passed into builder)
        option_cols = [
            c for c in df_virtual_devices.columns
            if c not in ("virtual_device", "physical_device")
        ]
        headers = desired_ids + sorted(option_cols)
        return pd.DataFrame(columns=headers)

    # --- VMware core‑count sanity fix -------------------------------------
    def _fix_vmware_row(r: pd.Series) -> pd.Series:
        """
        RVTools sometimes records 1 CPU / 1 core for ESXi hosts.  When
        virtualization_type == 'VMWare', parse the embedded `raw_data`
        (JSON string) and override bogus values.  Also capture ESXi version,
        cores‑per‑CPU, and vendor when available.
        """
        if str(r.get("virtualization_type")).lower() != "vmware":
            return r

        raw = r.get("raw_data")
        if not isinstance(raw, str) or not raw.strip():
            return r
        try:
            j = json.loads(raw)
        except Exception:
            return r

        to_int = lambda v, d=0: int(v) if str(v).isdigit() else d

        if r.get("total_number_of_processors", 0) <= 1:
            r["total_number_of_processors"] = to_int(j.get("# CPU"), 1)
        if r.get("total_number_of_cores", 0) <= 1:
            r["total_number_of_cores"] = to_int(j.get("# Cores"), 1)
        if r.get("total_number_of_threads", 0) <= 1:
            r["total_number_of_threads"] = to_int(j.get("# Cores"), 1) * 2

        if "Cores per CPU" in j:
            r["cores_per_cpu"] = to_int(j["Cores per CPU"])
        if "ESX Version" in j:
            r["esx_version"] = j["ESX Version"]
        if pd.isna(r.get("manufacturer")) and j.get("Vendor"):
            r["manufacturer"] = j["Vendor"]

        return r

    # apply the fix row‑wise
    df_host_specs = df_host_specs.apply(_fix_vmware_row, axis=1)

    # ── Explicit VMware sizing override from raw_data ─────────────────────────
    vm_mask = df_host_specs["virtualization_type"].str.lower() == "vmware"
    if vm_mask.any():
        def _override_vmware_sizing(r: pd.Series) -> pd.Series:
            raw = r.get("raw_data", "")
            try:
                payload = json.loads(raw) if raw else {}
            except Exception:
                import ast
                try:
                    payload = ast.literal_eval(raw) if raw else {}
                except Exception:
                    payload = {}
            # override processor/core values if present
            r["total_number_of_processors"] = payload.get("# CPU", r["total_number_of_processors"])
            r["cores_per_cpu"] = payload.get("Cores per CPU", r.get("cores_per_cpu", r["total_number_of_processors"]))
            r["total_number_of_cores"] = payload.get("# Cores", r["total_number_of_cores"])
            return r

        # apply override to VMware rows
        df_host_specs.loc[vm_mask] = df_host_specs[vm_mask].apply(_override_vmware_sizing, axis=1)

    id_cols_specs = [
        "device_name",               # host name (will match physical_device)
        "virtualization_type",
        "cluster_name",
        "operating_system_type",
        "esx_version",
        "total_number_of_processors",
        "cores_per_cpu",
        "total_number_of_cores",
        "total_number_of_threads",
        "cpu_model",
        "oracle_core_factor",
        "manufacturer",
        "model",
        "number_of_vms"              # vSphere’s own VM count
    ]
    df_host_specs = df_host_specs[id_cols_specs].drop_duplicates("device_name")
    df_host_specs.rename(columns={"device_name": "physical_device"}, inplace=True)

    # 2‑B  figure out which columns in df_virtual_devices are option symbols
    possible_sizing = [
        "number_of_virtual_threads", "cpu_threads", "lscpu_total_threads"
    ]
    sizing_cols_vm = [c for c in possible_sizing if c in df_virtual_devices.columns]

    # Ensure sizing columns are numeric (convert non-numeric or empty to 0)
    for col in sizing_cols_vm:
        df_virtual_devices[col] = pd.to_numeric(
            df_virtual_devices[col], errors="coerce"
        ).fillna(0).astype(int)

    option_cols = [
        c for c in df_virtual_devices.columns
        if c not in (
            ["virtual_device", "physical_device", "virtualization_type",
             "operating_system_type", "domain", "pool", "cpu_model"] + sizing_cols_vm
        )
    ]

    # 2‑C  convert symbols → priority for aggregation
    df_prior = df_virtual_devices.copy()
    for col in option_cols:
        df_prior[col] = df_prior[col].map(SYM2PRIO).fillna(0).astype(int)

    grouped = df_prior.groupby("physical_device", as_index=False)

    # max priority = highest tier of use
    df_opts_max = grouped[option_cols].max()

    # numeric sums for vCPU / vCore
    if sizing_cols_vm:
        df_sizing_sum = grouped[sizing_cols_vm].sum()
    else:
        df_sizing_sum = df_virtual_devices[["physical_device"]].drop_duplicates()
        df_sizing_sum["sizing_note"] = "NO VIRTUAL SIZING DATA FOUND"

    # stitch together
    df_hosts = (
        df_opts_max
        .merge(df_sizing_sum, on="physical_device", how="left")
        .merge(df_host_specs, on="physical_device", how="left")
    )

    # convert priorities back to symbols only for columns present
    for col in option_cols:
        if col in df_hosts.columns:
            df_hosts[col] = df_hosts[col].map(PRIO2SYM)

    # Ensure all expected physical sizing fields exist
    expected_physical_cols = [
        "total_number_of_processors", "total_number_of_cores", "total_number_of_threads"
    ]
    for col in expected_physical_cols:
        if col not in df_hosts.columns:
            df_hosts[col] = ""

    # ── drop redundant columns for hosts layout ────────────────────────────────
    df_hosts = df_hosts.drop(columns=[
        "total_number_of_threads", "capped", "cpu_speed",
        "device_manufacturer", "device_model", "hyper_threading",
        "lscpu_cores_per_socket", "lscpu_hypervisor",
        "lscpu_threads_per_core", "number_of_virtual_processors",
        "operating_system_release", "virtualization_type",
        "lscpu_total_threads", "oracle_core_factor_x", "oracle_core_factor_y"
    ], errors="ignore")

    # If split oracle_core_factor columns exist, prefer the host-spec one
    if "oracle_core_factor_x" in df_hosts.columns and "oracle_core_factor_y" in df_hosts.columns:
        df_hosts["oracle_core_factor"] = df_hosts["oracle_core_factor_y"].fillna(df_hosts["oracle_core_factor_x"])

    # tidy: place key host fields in specific order
    desired_ids = [
        "cluster_name",
        "physical_device",
        "number_of_vms",
        "operating_system_type",
        "esx_version",
        "manufacturer",
        "model",
        "cpu_model",
        "total_number_of_processors",
        "cores_per_cpu",
        "total_number_of_cores",
        "oracle_core_factor",
    ]
    id_block = [c for c in desired_ids if c in df_hosts.columns]

    # Recalculate option_block based on remaining columns
    option_block = [c for c in df_hosts.columns if c not in id_block]
    option_block.sort()
    final_cols = id_block + option_block
    df_hosts = df_hosts[final_cols].fillna("")

    return df_hosts


# ── 3. writer helper ────────────────────────────────────────────────────────
def write_hosts_sheet(
        wb: Workbook,
        df: pd.DataFrame,
        sheet_name: str = "Hosts"
) -> None:
    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])
    ws = wb.create_sheet(sheet_name)
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    ws.freeze_panes = ws["A2"]


# ── 4. convenience front‑door ───────────────────────────────────────────────
def build(
        df_virtual_devices: pd.DataFrame,
        physical_hosts_json: Path | str,
        workbook: Workbook | None = None
) -> tuple[Workbook, pd.DataFrame]:

    df_hosts = build_hosts_df(df_virtual_devices, physical_hosts_json)

    wb = workbook or Workbook()
    write_hosts_sheet(wb, df_hosts)

    return wb, df_hosts