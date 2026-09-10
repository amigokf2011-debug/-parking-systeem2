import os
import json
import datetime
from flask import Flask, render_template_string, request, jsonify
from apscheduler.schedulers.background import BackgroundScheduler

app = Flask(__name__)
DATA_FILE = "parking_data.json"
EXCEL_FILE = "遠雄車位管理.xlsx"

# 預設名單與車位
DEFAULT_USERS = ["陳惠娥", "吳品聰", "楊文明", "林惠婷", "許文蘭", "陳昶源", "王佑毓", "徐崇淵"]
DEFAULT_SPOTS = [
    {"floor": "B4", "id": "437", "owner": "陳惠娥", "rule": "週一～週五"},
    {"floor": "B4", "id": "452", "owner": "吳品聰", "rule": "週一～週五"},
    {"floor": "B4", "id": "466", "owner": "", "rule": ""},
    {"floor": "B5", "id": "578", "owner": "陳昶源", "rule": "週一～週三"},
    {"floor": "B5", "id": "512", "owner": "", "rule": ""},
    {"floor": "B5", "id": "596", "owner": "", "rule": ""},
    {"floor": "B5", "id": "533", "owner": "", "rule": ""},
    {"floor": "B5", "id": "581", "owner": "", "rule": ""},
    {"floor": "B6", "id": "602", "owner": "徐崇淵", "rule": "週一"},
    {"floor": "B6", "id": "612", "owner": "", "rule": ""},
    {"floor": "B6", "id": "634", "owner": "", "rule": ""},
    {"floor": "B6", "id": "621", "owner": "", "rule": ""}
]

def load_initial_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass

    users = list(DEFAULT_USERS)
    spots = list(DEFAULT_SPOTS)

    # 嘗試讀取 Excel（若無檔案則直接採用預設值，避免雲端崩潰）
    if os.path.exists(EXCEL_FILE):
        try:
            import pandas as pd
            df_users = pd.read_excel(EXCEL_FILE, sheet_name="使用人名單")
            if "使用人" in df_users.columns:
                users = df_users["使用人"].dropna().tolist()
            
            df_spots = pd.read_excel(EXCEL_FILE, sheet_name="車位管理")
            loaded_spots = []
            for _, row in df_spots.iterrows():
                loaded_spots.append({
                    "floor": str(row["樓層"]).strip(),
                    "id": str(row["車位編號"]).strip(),
                    "owner": "" if pd.isna(row["使用人"]) else str(row["使用人"]).strip(),
                    "rule": "" if pd.isna(row["保留區間"]) else str(row["保留區間"]).strip()
                })
            if loaded_spots:
                spots = loaded_spots
        except Exception as e:
            print(f"讀取 Excel 失敗，自動採用預設名單: {e}")

    data = {
        "users": users,
        "spots": spots,
        "today_records": {},
        "last_reset_date": str(datetime.date.today())
    }
    save_data(data)
    return data

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def daily_reset():
    data = load_initial_data()
    data["today_records"] = {}
    data["last_reset_date"] = str(datetime.date.today())
    save_data(data)
    print(f"[{datetime.datetime.now()}] 當日車位已自動重置！")

# 啟動定時排程（半夜 00:00 自動清空）
scheduler = BackgroundScheduler()
scheduler.add_job(daily_reset, "cron", hour=0, minute=0)
scheduler.start()

def is_rule_active_today(rule_str):
    if not rule_str:
        return False
    weekday = datetime.date.today().weekday()  # 0=週一, 1=週二 ... 6=週日
    if "週一～週五" in rule_str or "週一~週五" in rule_str:
        return 0 <= weekday <= 4
    elif "週一～週三" in rule_str or "週一~週三" in rule_str:
        return 0 <= weekday <= 2
    elif "週一" in rule_str:
        return weekday == 0
    return False

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>遠雄車位預約管理系統</title>
    <style>
        :root {
            --primary: #1e3a8a;
            --bg-color: #f1f5f9;
            --card-bg: #ffffff;
            --green: #10b981;
            --red: #ef4444;
            --yellow: #f59e0b;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
        body { background-color: var(--bg-color); color: #1e293b; padding-bottom: 30px; }
        .header {
            background: linear-gradient(135deg, #1e3a8a, #3b82f6);
            color: white;
            padding: 16px;
            text-align: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
            position: sticky;
            top: 0;
            z-index: 100;
        }
        .header h1 { font-size: 18px; font-weight: 700; }
        .header .date-info { font-size: 13px; opacity: 0.9; margin-top: 4px; }
        .legend-bar {
            display: flex;
            justify-content: center;
            gap: 12px;
            background: white;
            padding: 10px 8px;
            margin-bottom: 12px;
            font-size: 12px;
            border-bottom: 1px solid #e2e8f0;
        }
        .legend-item { display: flex; align-items: center; gap: 4px; }
        .legend-dot { width: 10px; height: 10px; border-radius: 50%; display: inline-block; }
        .dot-red { background: var(--red); }
        .dot-green { background: var(--green); }
        .dot-yellow { background: var(--yellow); }
        .container { max-width: 600px; margin: 0 auto; padding: 0 12px; }
        .floor-section {
            background: white;
            border-radius: 12px;
            padding: 12px;
            margin-bottom: 16px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.05);
        }
        .floor-title {
            font-size: 16px;
            font-weight: 800;
            color: var(--primary);
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            gap: 6px;
        }
        .parking-row {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(68px, 1fr));
            gap: 8px;
            justify-items: center;
        }
        .spot-card {
            width: 100%;
            height: 110px;
            border: 2.5px solid #cbd5e1;
            border-radius: 8px;
            background: #f8fafc;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            align-items: center;
            padding: 6px 4px;
            cursor: pointer;
            transition: all 0.2s;
        }
        .spot-card:active { transform: scale(0.96); }
        .spot-card.reserved { border-color: var(--red); background: #fef2f2; }
        .spot-card.available { border-color: var(--green); background: #f0fdf4; }
        .spot-card.booked { border-color: var(--yellow); background: #fffbeb; }
        .spot-id { font-size: 16px; font-weight: 800; color: #0f172a; }
        .spot-user { font-size: 12px; font-weight: 700; text-align: center; word-break: break-all; }
        .spot-status { font-size: 10px; padding: 2px 4px; border-radius: 4px; font-weight: 600; }
        .status-reserved { color: var(--red); background: #fee2e2; }
        .status-available { color: var(--green); background: #dcfce7; }
        .status-booked { color: #b45309; background: #fef3c7; }
        .btn-admin {
            display: block;
            width: 100%;
            text-align: center;
            padding: 12px;
            background: #475569;
            color: white;
            border-radius: 8px;
            text-decoration: none;
            font-size: 14px;
            font-weight: 600;
            margin-top: 20px;
            border: none;
            cursor: pointer;
        }
        .modal {
            display: none;
            position: fixed;
            top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(0,0,0,0.5);
            z-index: 200;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }
        .modal-content {
            background: white;
            border-radius: 16px;
            padding: 20px;
            width: 100%;
            max-width: 360px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.2);
        }
        .modal-title { font-size: 17px; font-weight: 700; margin-bottom: 12px; }
        .modal-select {
            width: 100%;
            padding: 10px;
            border: 1.5px solid #cbd5e1;
            border-radius: 8px;
            font-size: 15px;
            margin-bottom: 16px;
            background: white;
        }
        .btn-group { display: flex; gap: 8px; }
        .btn {
            flex: 1;
            padding: 10px;
            border-radius: 8px;
            border: none;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
        }
        .btn-primary { background: var(--primary); color: white; }
        .btn-danger { background: var(--red); color: white; }
        .btn-cancel { background: #e2e8f0; color: #475569; }
    </style>
</head>
<body>
    <div class="header">
        <h1>遠雄車位預約管理系統</h1>
        <div class="date-info" id="currentDate"></div>
    </div>

    <div class="legend-bar">
        <div class="legend-item"><span class="legend-dot dot-red"></span> 固定保留</div>
        <div class="legend-item"><span class="legend-dot dot-green"></span> 空位可登記</div>
        <div class="legend-item"><span class="legend-dot dot-yellow"></span> 今日已登記</div>
    </div>

    <div class="container" id="parkingContainer"></div>

    <div class="container">
        <button class="btn-admin" onclick="openAdminModal()">⚙️ 後台名單管理</button>
    </div>

    <div class="modal" id="actionModal">
        <div class="modal-content">
            <h3 class="modal-title" id="modalTitle">車位操作</h3>
            <div id="modalBody"></div>
        </div>
    </div>

    <div class="modal" id="adminModal">
        <div class="modal-content" style="max-height: 85vh; overflow-y: auto;">
            <h3 class="modal-title">後台名單管理</h3>
            <div style="display: flex; gap: 6px; margin-bottom: 12px;">
                <input type="text" id="newUserName" placeholder="輸入同仁姓名" style="flex:1; padding: 8px; border: 1px solid #cbd5e1; border-radius: 6px;">
                <button class="btn btn-primary" style="flex:0 0 60px;" onclick="addUser()">新增</button>
            </div>
            <div id="userListContainer" style="display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 16px;"></div>
            <button class="btn btn-cancel" style="width: 100%;" onclick="closeAdminModal()">關閉</button>
        </div>
    </div>

    <script>
        let appData = {};
        const weekdays = ["週一", "週二", "週三", "週四", "週五", "週六", "週日"];

        async function fetchData() {
            try {
                const res = await fetch('/api/status');
                appData = await res.json();
                renderUI();
            } catch (e) {
                console.error("載入失敗", e);
            }
        }

        function renderUI() {
            const today = new Date();
            const dateStr = `${today.getFullYear()}年${today.getMonth()+1}月${today.getDate()}日 (${weekdays[today.getDay() === 0 ? 6 : today.getDay() - 1]})`;
            document.getElementById('currentDate').innerText = `今天日期：${dateStr}`;

            const floors = ["B4", "B5", "B6"];
            const container = document.getElementById('parkingContainer');
            container.innerHTML = '';

            floors.forEach(floor => {
                const floorSpots = appData.spots.filter(s => s.floor === floor);
                if (floorSpots.length === 0) return;

                const section = document.createElement('div');
                section.className = 'floor-section';
                section.innerHTML = `<div class="floor-title">🚗 ${floor} 樓層停車區</div><div class="parking-row" id="row-${floor}"></div>`;
                container.appendChild(section);

                const row = section.querySelector(`#row-${floor}`);
                floorSpots.forEach(spot => {
                    const statusInfo = getSpotStatus(spot);
                    const card = document.createElement('div');
                    card.className = `spot-card ${statusInfo.className}`;
                    card.onclick = () => handleSpotClick(spot, statusInfo);
                    card.innerHTML = `
                        <div class="spot-id">${spot.id}</div>
                        <div class="spot-user">${statusInfo.userDisplay}</div>
                        <div class="spot-status ${statusInfo.badgeClass}">${statusInfo.statusText}</div>
                    `;
                    row.appendChild(card);
                });
            });
        }

        function getSpotStatus(spot) {
            const record = appData.today_records[spot.id];
            const isReservedDay = spot.is_active_today;

            if (record) {
                if (record.status === 'released') {
                    return { className: 'available', statusText: '已釋出(空)', userDisplay: '釋出空位', badgeClass: 'status-available', state: 'released' };
                } else if (record.status === 'booked') {
                    return { className: 'booked', statusText: '今日已登記', userDisplay: record.user, badgeClass: 'status-booked', state: 'booked' };
                }
            }

            if (isReservedDay && spot.owner) {
                return { className: 'reserved', statusText: '固定保留', userDisplay: spot.owner, badgeClass: 'status-reserved', state: 'reserved' };
            }

            return { className: 'available', statusText: '空位', userDisplay: '無人使用', badgeClass: 'status-available', state: 'empty' };
        }

        function handleSpotClick(spot, statusInfo) {
            const modal = document.getElementById('actionModal');
            const title = document.getElementById('modalTitle');
            const body = document.getElementById('modalBody');
            title.innerText = `車位 ${spot.floor} - ${spot.id}`;

            let content = '';
            const userOptions = appData.users.map(u => `<option value="${u}">${u}</option>`).join('');

            if (statusInfo.state === 'reserved') {
                content = `
                    <p style="margin-bottom: 12px; font-size: 14px;">此車位目前固定保留給 <b>${spot.owner}</b>。</p>
                    <p style="margin-bottom: 16px; font-size: 13px; color: #64748b;">若今日不開車，可釋出給同仁登記：</p>
                    <div class="btn-group">
                        <button class="btn btn-danger" onclick="releaseSpot('${spot.id}')">今日釋出車位</button>
                        <button class="btn btn-cancel" onclick="closeModal()">取消</button>
                    </div>
                `;
            } else if (statusInfo.state === 'released') {
                content = `
                    <p style="margin-bottom: 12px; font-size: 14px; color: var(--green);">此車位已由原保留人釋出！</p>
                    <label style="font-size: 13px; font-weight:600; display:block; margin-bottom:4px;">請選擇登記人：</label>
                    <select id="userSelect" class="modal-select">${userOptions}</select>
                    <div class="btn-group" style="margin-bottom: 8px;">
                        <button class="btn btn-primary" onclick="bookSpot('${spot.id}')">登記此車位</button>
                        <button class="btn btn-cancel" onclick="closeModal()">返回</button>
                    </div>
                    <button class="btn btn-cancel" style="width: 100%; font-size: 12px;" onclick="cancelBooking('${spot.id}')">我是原保留人（收回車位）</button>
                `;
            } else if (statusInfo.state === 'empty') {
                content = `
                    <p style="margin-bottom: 12px; font-size: 14px;">此車位目前為空位，請選擇同仁姓名登記：</p>
                    <select id="userSelect" class="modal-select">${userOptions}</select>
                    <div class="btn-group">
                        <button class="btn btn-primary" onclick="bookSpot('${spot.id}')">確認登記</button>
                        <button class="btn btn-cancel" onclick="closeModal()">取消</button>
                    </div>
                `;
            } else if (statusInfo.state === 'booked') {
                content = `
                    <p style="margin-bottom: 12px; font-size: 14px;">目前已由 <b>${statusInfo.userDisplay}</b> 登記使用。</p>
                    <div class="btn-group">
                        <button class="btn btn-danger" onclick="cancelBooking('${spot.id}')">取消此登記</button>
                        <button class="btn btn-cancel" onclick="closeModal()">返回</button>
                    </div>
                `;
            }

            body.innerHTML = content;
            modal.style.display = 'flex';
        }

        function closeModal() { document.getElementById('actionModal').style.display = 'none'; }
        function openAdminModal() {
            renderUserList();
            document.getElementById('adminModal').style.display = 'flex';
        }
        function closeAdminModal() { document.getElementById('adminModal').style.display = 'none'; }

        function renderUserList() {
            const container = document.getElementById('userListContainer');
            container.innerHTML = appData.users.map(u => `
                <span style="background: #e2e8f0; padding: 4px 8px; border-radius: 4px; font-size: 13px; display: inline-flex; align-items: center; gap: 4px;">
                    ${u} <span style="cursor: pointer; color: red;" onclick="deleteUser('${u}')">×</span>
                </span>
            `).join('');
        }

        async function bookSpot(spotId) {
            const user = document.getElementById('userSelect').value;
            await fetch('/api/book', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ spot_id: spotId, user: user })
            });
            closeModal();
            fetchData();
        }

        async function releaseSpot(spotId) {
            await fetch('/api/release', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ spot_id: spotId })
            });
            closeModal();
            fetchData();
        }

        async function cancelBooking(spotId) {
            await fetch('/api/cancel', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ spot_id: spotId })
            });
            closeModal();
            fetchData();
        }

        async function addUser() {
            const name = document.getElementById('newUserName').value.trim();
            if (!name) return;
            await fetch('/api/users', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ name: name })
            });
            document.getElementById('newUserName').value = '';
            fetchData();
            renderUserList();
        }

        async function deleteUser(name) {
            if (!confirm(`確定要刪除 ${name} 嗎？`)) return;
            await fetch('/api/users/delete', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ name: name })
            });
            fetchData();
            renderUserList();
        }

        fetchData();
    </script>
</body>
</html>
"""

@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route("/api/status")
def api_status():
    data = load_initial_data()
    if data.get("last_reset_date") != str(datetime.date.today()):
        daily_reset()
        data = load_initial_data()

    enriched_spots = []
    for s in data["spots"]:
        s_copy = dict(s)
        s_copy["is_active_today"] = is_rule_active_today(s.get("rule", ""))
        enriched_spots.append(s_copy)

    return jsonify({
        "users": data["users"],
        "spots": enriched_spots,
        "today_records": data.get("today_records", {})
    })

@app.route("/api/book", methods=["POST"])
def api_book():
    req = request.json or {}
    spot_id = str(req.get("spot_id"))
    user = req.get("user")
    data = load_initial_data()
    data["today_records"][spot_id] = {"status": "booked", "user": user}
    save_data(data)
    return jsonify({"success": True})

@app.route("/api/release", methods=["POST"])
def api_release():
    req = request.json or {}
    spot_id = str(req.get("spot_id"))
    data = load_initial_data()
    data["today_records"][spot_id] = {"status": "released", "user": ""}
    save_data(data)
    return jsonify({"success": True})

@app.route("/api/cancel", methods=["POST"])
def api_cancel():
    req = request.json or {}
    spot_id = str(req.get("spot_id"))
    data = load_initial_data()
    if spot_id in data["today_records"]:
        del data["today_records"][spot_id]
        save_data(data)
    return jsonify({"success": True})

@app.route("/api/users", methods=["POST"])
def api_add_user():
    req = request.json or {}
    name = req.get("name", "").strip()
    data = load_initial_data()
    if name and name not in data["users"]:
        data["users"].append(name)
        save_data(data)
    return jsonify({"success": True})

@app.route("/api/users/delete", methods=["POST"])
def api_delete_user():
    req = request.json or {}
    name = req.get("name", "").strip()
    data = load_initial_data()
    if name in data["users"]:
        data["users"].remove(name)
        save_data(data)
    return jsonify({"success": True})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)