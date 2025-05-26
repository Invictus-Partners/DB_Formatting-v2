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
        # Physical Hosts.json missing or invalid: try to load declaration CSV fallback
        declaration_csv = Path("input/host_declaration_template.csv")
        if declaration_csv.exists():
            print("⚠️  Physical host JSON missing — using declaration CSV instead")
            df_host_specs = pd.read_csv(declaration_csv)
            # Normalize column headers: lowercase and strip whitespace
            df_host_specs.columns = df_host_specs.columns.str.strip().str.lower()

            # Rename using normalized column names
            df_host_specs.rename(columns={
                "virtual device": "virtual_device",
                "physical host": "physical_device",
                "model": "model_y",
                "manufacturer": "manufacturer",
                "cpu model": "cpu_model",
                "# processors": "total_number_of_processors",
                "cores per cpu": "cores_per_cpu",
                "total cores": "total_number_of_cores"
            }, inplace=True)
            df_host_specs["device_name"] = df_host_specs["physical_device"]
            df_host_specs["model"] = df_host_specs["model_y"]
            df_host_specs["virtualization_type"] = "declared"
            df_host_specs["cluster_name"] = "DECLARED"
            df_host_specs["operating_system_type"] = "unknown"
            df_host_specs["esx_version"] = "N/A"
            df_host_specs["data_source"] = "declared"
            df_host_specs["model"] = df_host_specs.get("model_y", "")

            # Merge declared physical_device values into virtual_devices
            df_virtual_devices = df_virtual_devices.merge(
                df_host_specs[["virtual_device", "physical_device"]],
                on="virtual_device",
                how="left"
            )

            # Compute number_of_vms as count of VMs per host
            vm_counts = df_virtual_devices.groupby("physical_device")["virtual_device"].count().reset_index()
            vm_counts.rename(columns={"virtual_device": "number_of_vms"}, inplace=True)
            df_host_specs = df_host_specs.drop(columns=["number_of_vms"], errors="ignore").merge(vm_counts, on="physical_device", how="left")

            if df_virtual_devices["physical_device"].isna().any():
                print("⚠️ Some virtual devices did not map to a declared physical host")
            # Ensure all expected host spec columns exist when using declaration CSV
            for col in [
                "device_name", "virtualization_type", "cluster_name",
                "operating_system_type", "esx_version",
                "total_number_of_processors", "cores_per_cpu", "total_number_of_cores",
                "total_number_of_threads", "cpu_model", "oracle_core_factor",
                "manufacturer", "model", "number_of_vms"
            ]:
                if col not in df_host_specs.columns:
                    df_host_specs[col] = ""
        else:
            # Return empty host sheet with standard columns
            print("⚠️ No physical host data or declaration file found – this is required for final calculation!!!")
            desired_ids = [
                "cluster_name", "physical_device", "number_of_vms",
                "operating_system_type", "esx_version", "manufacturer",
                "model", "cpu_model", "total_number_of_processors",
                "cores_per_cpu", "total_number_of_cores", "oracle_core_factor",
            ]
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

    excluded_cols = (
        ["virtual_device", "physical_device", "virtualization_type",
         "operating_system_type", "domain", "pool", "cpu_model",
         "number_of_virtual_processors", "number_of_virtual_cores", "number_of_virtual_threads"] + sizing_cols_vm
    )

    option_cols = [c for c in df_virtual_devices.columns if c not in excluded_cols]

    # 2‑C  convert symbols → priority for aggregation
    df_prior = df_virtual_devices.copy()
    for col in option_cols:
        if col in df_prior.columns and isinstance(SYM2PRIO, dict):
            try:
                df_prior[col] = df_prior[col].astype(str).map(SYM2PRIO).fillna(0).astype(int)
            except Exception:
                print(f"⚠️ Failed mapping column '{col}', assigning 0")
                df_prior[col] = 0
        else:
            df_prior[col] = 0

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

    if "cluster_name" in df_host_specs.columns:
        df_hosts["data_source"] = df_hosts["physical_device"].map(
            df_host_specs.set_index("physical_device")["cluster_name"]
        ).fillna("discovered").apply(lambda x: "declared" if x == "DECLARED" else "discovered")
    else:
        df_hosts["data_source"] = "discovered"

    # convert priorities back to symbols only for columns present
    for col in option_cols:
        if (
            col in df_hosts.columns
            and df_hosts[col].dtype in ["int64", "float64"]
            and isinstance(PRIO2SYM, dict)
        ):
            try:
                df_hosts[col] = df_hosts[col].map(PRIO2SYM)
            except Exception:
                print(f"⚠️ Failed converting priority to symbol for '{col}'")

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

    # Preserve declared model names from declaration CSV if model missing or all NaN
    if "model" not in df_hosts.columns or df_hosts["model"].isna().all():
        df_hosts["model"] = df_hosts.get("model_y", "")

    # tidy: place key host fields in specific order
    desired_ids = [
        "cluster_name",
        "physical_device",
        "number_of_vms",
        "operating_system_type",
        "esx_version",
        "model_x",
        "model_y",
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
    ws = wb[sheet_name]
    # Write DataFrame including header starting at cell B3
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=3):
        for c_idx, value in enumerate(row, start=2):  # Column B = 2
            ws.cell(row=r_idx, column=c_idx, value=value)


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