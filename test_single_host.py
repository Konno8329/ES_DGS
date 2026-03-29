import time
import re
import pandas as pd
from netmiko import ConnectHandler
from netmiko.exceptions import NetmikoTimeoutException, NetmikoAuthenticationException

# テスト用に1台だけを対象にします
HOSTS = [
    {"hostname": "該当ホスト名", "ip": "該当IPアドレス"}
]

OFFICIAL_MODULES = [
    "DEM-431XT", "DEM-432XT", "DEM-410T", # 10G
    "DEM-311GT", "DEM-310GT", "DEM-330T", "DEM-330R", "DGS-712" # 1G
]

def get_password(hostname):
    """ホスト名の先頭2文字からパスワードを判定"""
    if hostname.startswith("ES"):
        return "該当パスワード(ESシリーズ)"
    elif hostname.startswith("BS"):
        return "該当パスワード(BSシリーズ)"
    return ""

def determine_connection_type(port_type, vendor_pn, compliance):
    if "1000BASE-T" in port_type or "1000BASE-T" in compliance or vendor_pn in ["DEM-410T", "DGS-712"]:
        return "銅 (Copper)"
    if "reading" in vendor_pn.lower():
        return "不明 (reading)"

    comp_upper = compliance.upper()
    if "SX" in comp_upper or "SR" in comp_upper or vendor_pn in ["DEM-311GT", "DEM-431XT"]:
        return "マルチ光 (MMF)"
    if "LX" in comp_upper or "LR" in comp_upper or "BX" in comp_upper or vendor_pn in ["DEM-310GT", "DEM-432XT", "DEM-330T", "DEM-330R"]:
        return "シングル光 (SMF)"

    if "LC" in port_type or "SFP" in port_type:
        return "光 (詳細不明)"

    return "-"

def collect_switch_data(host_info):
    hostname = host_info["hostname"]
    ip = host_info["ip"]
    password = get_password(hostname)

    print(f"\n[{hostname}] ({ip}) へのテスト接続を開始します...")

    log_filename = f"{hostname}_teraterm_log.txt"
    device = {
        'device_type': 'cisco_ios',
        'host': ip,
        'username': '該当ユーザー名',
        'password': password,
        'global_delay_factor': 2,
        'session_log': log_filename
    }

    port_data = []

    try:
        with ConnectHandler(**device) as net_connect:
            net_connect.send_command("terminal length 0")

            desc_output = net_connect.send_command("show interfaces description")
            descriptions = {}
            for line in desc_output.splitlines():
                parts = line.split()
                if len(parts) >= 2 and parts[0].startswith("eth"):
                    desc_text = " ".join(parts[3:]) if len(parts) > 3 else ""
                    descriptions[parts[0]] = desc_text

            status_output = net_connect.send_command("show interfaces status")

            for line in status_output.splitlines():
                parts = line.split()
                if not parts or not parts[0].startswith("eth"):
                    continue

                port = parts[0] # 例: eth1/0/26 や eth1/0/21(c)
                status = parts[1] if len(parts) > 1 else ""
                duplex = parts[3] if len(parts) > 3 else ""
                speed = parts[4] if len(parts) > 4 else ""
                port_type = parts[5] if len(parts) > 5 else ""
                desc = descriptions.get(port, "")

                vendor_pn = "-"
                compliance = "-"

                # 🌟【修正箇所】コマンド用にポート番号の文字列を綺麗に整形する
                # "eth1/0/26" -> "1/0/26", "eth1/0/21(c)" -> "1/0/21"
                port_id_for_cmd = port.replace("eth", "").replace("(c)", "").replace("(f)", "")

                if status == "connected" and port_type != "1000BASE-T" and "(c)" not in port:
                    # 修正したクリーンなポート番号を使ってコマンドを送信
                    gbic_cmd = f"show interfaces ethernet {port_id_for_cmd} gbic"
                    gbic_output = net_connect.send_command(gbic_cmd)

                    print(f"\n--- 取得結果: {port} ---")
                    print(gbic_output.strip())
                    print("--------------------------")

                    if "reading..." in gbic_output:
                        print(f"  -> {port}: reading... を検知。1秒待機して再取得します...")
                        time.sleep(1)
                        gbic_output = net_connect.send_command(gbic_cmd)

                    if "reading..." in gbic_output:
                        vendor_pn = "不明 (reading...)"
                    else:
                        pn_match = re.search(r"Vendor PN:\s*(.+)", gbic_output)
                        comp_match = re.search(r"Ethernet Compliance Code:\s*(.+)", gbic_output)

                        if pn_match:
                            vendor_pn = pn_match.group(1).strip()
                        if comp_match:
                            compliance = comp_match.group(1).strip()

                elif status != "connected":
                    vendor_pn = "未接続"
                elif "(c)" in port or port_type == "1000BASE-T":
                    vendor_pn = "(RJ-45 内蔵)"
                    compliance = "1000BASE-T"

                conn_type = determine_connection_type(port_type, vendor_pn, compliance)

                port_data.append({
                    "ポート番号": port,
                    "リンク状態": status,
                    "接続方式": conn_type,
                    "モジュール型番": vendor_pn,
                    "規格": compliance,
                    "速度 (Speed)": speed,
                    "Duplex": duplex,
                    "Type": port_type,
                    "接続先 (Description)": desc
                })

        print(f"[{hostname}] データの取得が完了しました。")
        return hostname, port_data

    except Exception as e:
        print(f"[{hostname}] エラーが発生しました: {e}")
        return hostname, []

def main():
    all_data = {}
    summary_data = []

    for host in HOSTS:
        hostname, data = collect_switch_data(host)
        if data:
            all_data[hostname] = data

            counts = {mod: 0 for mod in OFFICIAL_MODULES}
            counts["不明・非純正 (reading...)"] = 0
            counts["備え付け銅線 (コンボ等)"] = 0
            counts["空きスロット (未接続/ダウン)"] = 0

            for row in data:
                pn = row["モジュール型番"]
                status = row["リンク状態"]

                if status != "connected":
                    counts["空きスロット (未接続/ダウン)"] += 1
                elif "reading" in pn:
                    counts["不明・非純正 (reading...)"] += 1
                elif "RJ-45" in pn or row["接続方式"] == "銅 (Copper)":
                    counts["備え付け銅線 (コンボ等)"] += 1
                else:
                    matched = False
                    for mod in OFFICIAL_MODULES:
                        if mod in pn:
                            counts[mod] += 1
                            matched = True
                            break
                    if not matched:
                        counts[pn] = counts.get(pn, 0) + 1

            summary_row = {"ホスト名": hostname}
            summary_row.update(counts)
            summary_data.append(summary_row)

    if not all_data:
        print("データが取得できませんでした。")
        return

    output_file = "TEST_Single_Host_Inventory.xlsx"
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_summary = pd.DataFrame(summary_data)
        df_summary = df_summary.loc[:, (df_summary != 0).any(axis=0)]
        df_summary.to_excel(writer, sheet_name="00_総合集計", index=False)

        for hostname, data in all_data.items():
            safe_sheet_name = hostname[:31]
            df_details = pd.DataFrame(data)
            df_details.to_excel(writer, sheet_name=safe_sheet_name, index=False)

            worksheet = writer.sheets[safe_sheet_name]
            worksheet.auto_filter.ref = worksheet.dimensions

    print(f"\n完了しました！ '{output_file}' に結果を出力しました。")

if __name__ == "__main__":
    main()
