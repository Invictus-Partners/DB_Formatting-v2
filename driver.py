from builders import evidence_tab, db_details_tab, virtual_devices_tab, hosts_tab
from builders.clusters_tab import build as build_clusters
import os
from datetime import datetime

# 1) Evidence …
wb = evidence_tab.build_evidence_tab("input/Options Usage Evidence.json")

# 2) DB Details …
wb, df_db = db_details_tab.build(
    "input/Oracle Databases Installed.json",
    "input/Options Usage Evidence.json",
    workbook=wb
)

# 2a) determine which VMs host databases
allowed_vms = df_db["device_name"].unique().tolist()

# 3) Virtual Devices …
wb, df_vm = virtual_devices_tab.build(
    "input/Virtual Devices.json",
    allowed_vms=allowed_vms,
    df_db_details=df_db,
    workbook=wb
)

# 4) Hosts …
wb, df_hosts = hosts_tab.build(
    df_vm,
    "input/Physical Hosts.json",
    workbook=wb
)

# 5) Clusters – roll-up per cluster
wb, df_clusters = build_clusters(
    df_hosts,
    "input/Virtualization Clusters.json",
    workbook=wb
)

if __name__ == "__main__":
    # ensure the output directory exists
    output_dir = "final"
    os.makedirs(output_dir, exist_ok=True)

    # timestamp the filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"oracle_audit_rollup_{timestamp}.xlsx"
    output_path = os.path.join(output_dir, output_file)

    wb.save(output_path)
    print(f"✓ Roll-up complete → {output_path}")