# DB Formatting v2

A small Python-based tool to collect, transform and consolidate Oracle database & infrastructure data into a single Excel workbook.

---

## 📂 Repository layout

```
DB_Formatting_v2/
├─ builders/
│  ├─ db_details_tab.py       # builds “Database Details” tab
│  ├─ evidence_tab.py         # builds “Evidence” tab
│  ├─ virtual_devices_tab.py  # builds “Virtual Devices” tab
│  ├─ hosts_tab.py            # builds “Hosts” tab
│  └─ clusters_tab.py         # builds “Clusters” tab
├─ input/
│  ├─ Databases on Infrastructure.json       # DB instance metadata
│  ├─ Options Usage Evidence.json            # per-DB option usage records
│  ├─ Virtual Devices.json                   # VM inventory + raw_data blob
│  ├─ Physical Hosts.json                    # Host inventory + raw_data blob
│  └─ Virtualization Clusters.json           # cluster-level metadata & raw_data blob
├─ driver.py                 # orchestrates all tabs & writes Excel
├─ oracle_audit_rollup.xlsx  # (auto-generated) final output
└─ requirements.txt
```

---

## 🔑 What files you need

Place these five JSON exports (exact filenames matter!) into `./input` before running:

1. **Databases on Infrastructure.json**  
   - One row per database instance.  
   - Fields: `device_name`, `database_name`, `product_version`, etc.

2. **Options Usage Evidence.json**  
   - One row per option-usage event: `(device_name, database_name, option_name, result)`.  
   - Used to derive “+ / ~ / # / ^” symbols per option, per database.

3. **Virtual Devices.json**  
   - One row per VM.  
   - Fields: `device_name`, `physical_host`, OS, CPU, plus a `raw_data` JSON blob.

4. **Physical Hosts.json**  
   - One row per physical host.  
   - Fields: `device_name` (host), cluster, CPU/socket count, plus `raw_data`.

5. **Virtualization Clusters.json**  
   - One row per cluster.  
   - Fields: `device_name` (cluster name), `number_of_cluster_members`, `total_number_of_processors`, `total_number_of_cores`, `total_number_of_threads`, plus a `raw_data` JSON blob with cluster settings.

> **Note:** if **Physical Hosts.json** is missing or invalid, the tool will still build an **empty** “Hosts” sheet (with headers only), so you can review VM-level data without failure.

---

## ⚙️ Installation

1. Clone this repo:  
   ```bash
   git clone <repo-url>
   cd DB_Formatting_v2
   ```

2. Create & activate a Python venv, install deps:  
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt
   ```

---

## ▶️ Running the tool

From the repo root:

```bash
python driver.py
```

This will:

1. **Evidence** tab: import raw option-usage evidence.  
2. **Database Details** tab: merge DB metadata + pivot option symbols.  
3. **Virtual Devices** tab: join VM inventory + DB-derived option symbols.  
4. **Hosts** tab: roll up VM symbols & sizing to each physical host.  
5. **Clusters** tab: roll up host data into one row per cluster using Virtualization Clusters.json.

When complete, opens (or updates) `oracle_audit_rollup.xlsx` with four worksheets.

---

## 📝 Worksheet summaries

- **Evidence**  
  Lists every raw evidence row from `Options Usage Evidence.json`.

- **Database Details**  
  One row per database instance, with DB metadata plus one symbol-column per Oracle option (`+`, `~`, `#`, `^`).

- **Virtual Devices**  
  One row per VM (filtered to only those hosting databases).  
  Columns include host name, OS, CPU specs (unpacked from `raw_data`), **plus** the highest-priority option symbol for each VM.

- **Hosts**  
  One row per physical host.  
  Aggregates across its VMs: sums CPU sockets, cores, and picks the highest-priority symbol for each option.

- **Clusters**
  One row per cluster. Columns include:
  - Cluster identity: `cluster_name`, `number_of_hosts`, `instance_uuid`, `visdk_server`
  - Cluster health & configuration: `ha_enabled`, `drs_enabled`, `admission_control_enabled`, `num_vmotions`, `overall_status`
  - Sizing: `total_number_of_processors`, `total_number_of_cores`
  - Merged option usage symbols (one symbol column per Oracle option)

---

## ⚙️ Customisation & Troubleshooting

- **Filtering VMs**  
  By default VMs are filtered to only those in your Database Details sheet. You can adjust or remove this filter in `driver.py`.

- **Missing Hosts**  
  If `Physical Hosts.json` is not present or is malformed, the Hosts tab will be created empty (headers only) instead of erroring out.

- **Raw_data overrides**  
  VMware hosts extract sizing (`# CPU`, `Cores per CPU`, `# Cores`) directly from their `raw_data` blob to avoid inconsistencies.

- **Column ordering**  
  You can tweak the `id_block` in each builder (e.g. `virtual_devices_tab.py`, `hosts_tab.py`) to reorder or remove columns to suit your reporting needs.


> _Built by the AutoGen team for Oracle audit roll-up._