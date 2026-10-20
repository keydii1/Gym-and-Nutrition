# 🏛️ HỆ THỐNG QUẢN TRỊ TĂNG CÂN & TĂNG CƠ TRÊN NOTION (BULKING OS 2026)
> **Thiết kế tối ưu hóa cho Ho Hoàng Sơn** • Thể trạng khởi điểm: 50.3 kg / 175 cm (BMI 16.4) • Mục tiêu dài hạn: 65.0 kg (+14.7 kg) • Tinh gọn, thẩm mỹ (Aesthetic Dark/Minimalist), không ngợp dữ liệu như Excel.

---

## 🌟 TẠI SAO NOTION TỐT HƠN EXCEL CHO MỤC TIÊU CỦA BẠN?

1. **Giao diện sạch sẽ, trực quan (Aesthetic Dashboard)**: Không còn hàng trăm ô số san sát gây mỏi mắt. Dữ liệu được hiển thị dưới dạng **Thẻ ảnh (Gallery)**, **Bảng Kanban (Board)**, hoặc **Danh sách tinh gọn (List)**.
2. **Nhập liệu 1 chạm trên điện thoại (Mobile Friendly)**: Có các nút bấm **Quick Action Buttons** (`+ Cân sáng`, `+ Log bữa ăn`, `+ Log bài tập`) ngay đầu trang chủ, mở app là bấm nhập được ngay trong 10 giây tại phòng gym hoặc quán ăn.
3. **Phân tách rạch ròi 2 Hub độc lập**:
   - 🏋️ **Gym Hub**: Chỉ tập trung vào giáo án, bài tập, số tạ và hiệp tập.
   - 🥗 **Nutrition Hub**: Chỉ tập trung vào món ăn, năng lượng và công thức nấu nướng cho sinh viên.
4. **Liên kết thông minh (Relations & Rollups)**: Chọn món ăn tự động liên kết với thực phẩm; chọn bài tập tự động hiện video hướng dẫn và mẹo an toàn.

---

## 📂 DANH SÁCH FILE DATABASE ĐÃ SẴN SÀNG IMPORT (Trong thư mục `notion_templates/`)

Hệ thống đã tự động xuất sẵn **8 file CSV chuẩn hóa UTF-8** để bạn import trực tiếp vào Notion:

| File CSV | Hub tương ứng | Dữ liệu bao gồm |
|---|---|---|
| `01_Gym_Exercise_Library.csv` | 🏋️ Gym Hub | 35 bài tập chuẩn Hypertrophy, phân loại nhóm cơ, dụng cụ, mẹo kỹ thuật, bài thay thế |
| `02_Gym_Workout_Routines.csv` | 🏋️ Gym Hub | Giáo án Upper/Lower 4 buổi/tuần, số hiệp, reps, RIR và thời gian nghỉ |
| `03_Gym_Workout_Logs_Template.csv` | 🏋️ Gym Hub | Mẫu nhật ký set tập (Tạ, Reps, RIR, Cờ đau khớp, Volume) |
| `04_Nutrition_Food_Database.csv` | 🥗 Nutrition Hub | 42 thực phẩm thông dụng Việt Nam (Kcal, P, C, F, Giá, Sống/Chín rõ ràng) |
| `05_Nutrition_Recipes_Database.csv` | 🥗 Nutrition Hub | 6 công thức sinh tố mass và cơm tăng cân sinh viên (Kcal/suất, P/suất, cách làm) |
| `06_Nutrition_Daily_Meal_Log_Template.csv` | 🥗 Nutrition Hub | Mẫu nhật ký ăn uống theo từng bữa (Sáng, Trưa, Xế, Tối) |
| `07_Body_Weight_Tracker.csv` | 📈 Progress Hub | Dữ liệu theo dõi cân nặng sáng, giờ ngủ, số đo vòng eo, ngực, tay |
| `08_InBody_History.csv` | 📈 Progress Hub | Lịch sử InBody (Bản ghi baseline 19/10/2026: 50.3kg, SMM 25.7kg, PBF 7.8%) |

---

## 🛠️ HƯỚNG DẪN 3 BƯỚC KHỞI TẠO WORKSPACE TRÊN NOTION (MẤT ~5 PHÚT)

### Bước 1: Tạo một trang trắng mới (Parent Page)
1. Mở Notion $\rightarrow$ Tạo trang mới, đặt tên: `⚡ BULKING & HYPERTROPHY OS 2026`.
2. Chọn biểu tượng icon: `⚡` hoặc `🏋️`.
3. Thêm ảnh bìa (Cover Image): Tìm từ khóa Unsplash `minimalist gym` hoặc `fitness aesthetic dark`.
4. Bật chế độ **Full Width** (bấm vào dấu `...` góc phải trên cùng $\rightarrow$ bật `Full width`).

### Bước 2: Import 8 Database từ thư mục `notion_templates/`
1. Trong trang vừa tạo, gõ `/import` hoặc bấm vào nút `...` $\rightarrow$ chọn **Import**.
2. Chọn định dạng **CSV**.
3. Lần lượt tải lên 8 file CSV trong thư mục `notion_templates/` của bạn:
   - `01_Gym_Exercise_Library.csv`
   - `02_Gym_Workout_Routines.csv`
   - `03_Gym_Workout_Logs_Template.csv`
   - `04_Nutrition_Food_Database.csv`
   - `05_Nutrition_Recipes_Database.csv`
   - `06_Nutrition_Daily_Meal_Log_Template.csv`
   - `07_Body_Weight_Tracker.csv`
   - `08_InBody_History.csv`
4. Notion sẽ tự động chuyển mỗi file thành một Database hoàn chỉnh với đầy đủ các cột dữ liệu. Bạn kéo 8 database này vào một mục con gọi là `🗄️ Master Databases (Lưu trữ)`.

### Bước 3: Thiết lập Dashboard 2 Cột Đẹp Mắt
Tạo giao diện trang chủ theo sơ đồ wireframe dưới đây.

---

## 📐 SƠ ĐỒ WIREFRAME TRANG CHỦ (MASTER DASHBOARD LAYOUT)

```text
+--------------------------------------------------------------------------------------------------+
| ⚡ BULKING & HYPERTROPHY OS (HO HOANG SON)                                                       |
| 🎯 Cân nặng: 50.3 kg -> 65.0 kg (+14.7 kg) | 🥗 Năng lượng: 2.100 kcal | 🏋️ Lịch: 4 buổi/tuần   |
+--------------------------------------------------------------------------------------------------+
| [⚡ + Cân Sáng Hôm Nay]        [🥗 + Log Bữa Ăn Mới]         [🏋️ + Ghi Set Tập Hôm Nay]          |
+--------------------------------------------------------------------------------------------------+
|                                                                                                  |
| 🏋️ CỘT TRÁI: GYM HUB                                 🥗 CỘT PHẢI: NUTRITION HUB                 |
|                                                                                                  |
| 1. LỊCH TẬP HÔM NAY (Filtered Today)                 1. NHẬT KÝ ĂN UỐNG HÔM NAY (Board View)     |
|    - Upper A: Ngực & Lưng Xô                         - [Bữa Sáng] Sandwich trứng bơ đậu phộng    |
|    - 12 working sets • RIR 1-2 • Nghỉ 90-120s        - [Bữa Phụ]  Chuối tiêu + Sữa tươi         |
|    - [Checklist hoàn thành từng bài]                  - [Bữa Trưa] Cơm trắng + Ức gà áp chảo     |
|                                                      - [Bữa Tối]  Cơm trắng + Thịt nạc xào tofu |
| 2. DANH MỤC BÀI TẬP (Gallery Cards)                                                              |
|    - Thẻ hình ảnh ngực, lưng, vai, chân              2. MEAL-PREP MÓN TĂNG CÂN (Gallery Cards)  |
|    - Bấm vào xem video & mẹo an toàn                 - Sinh tố Mass tự làm (625 kcal)            |
|                                                      - Yến mạch ngâm sữa chua (430 kcal)        |
| 3. TIẾN BỘ TẠ & REPS (Progressive Overload)                                                      |
|    - Incline Dumbbell Press: 8kg x 10 reps           3. TRA CỨU THỰC PHẨM NHANH (Table View)    |
|    - Lat Pulldown: 30kg x 10 reps                    - Lọc theo Đạm / Tinh bột / Chất béo tốt   |
|                                                                                                  |
+--------------------------------------------------------------------------------------------------+
| 📈 PHẦN CUỐI TRANG: THEO DÕI TIẾN ĐỘ & INBODY                                                   |
| - Biểu đồ cân nặng 7 ngày gần nhất (Average Trend)                                               |
| - Thanh tiến độ visual về đích mốc 65 kg (Formula 2.0 Progress Bar)                              |
| - Lịch sử InBody định kỳ (4 - 8 tuần đo 1 lần)                                                  |
+--------------------------------------------------------------------------------------------------+
```

---

## 💡 CÔNG THỨC NOTION FORMULA 2.0 SIÊU ĐẸP (COPY & PASTE NGAY)

Notion vừa nâng cấp hệ thống Formula 2.0 rất mạnh mẽ. Dưới đây là các công thức được tối ưu riêng cho tài khoản của bạn:

### 1. Thanh tiến độ Visual về đích 65 kg (Visual Progress Bar)
Thêm một thuộc tính kiểu `Formula` vào database `07_Body_Weight_Tracker`, đặt tên là `Progress to 65kg`:
```javascript
let(
  start, 50.3,
  goal, 65.0,
  current, prop("Morning Weight (kg)"),
  pct, if(current > start, (current - start) / (goal - start), 0),
  pctClamped, min(max(pct, 0), 1),
  filled, round(pctClamped * 10),
  empty, 10 - filled,
  slice("▓▓▓▓▓▓▓▓▓▓", 0, filled) + slice("░░░░░░░░░░", 0, empty) + " " + round(pctClamped * 100) + "%"
)
```
*Kết quả hiển thị*: `▓▓░░░░░░░░ 15%` hoặc `▓▓▓▓▓░░░░░ 50%` rất trực quan và sang trọng!

### 2. Tự động tính BMI và Cảnh báo Thể trạng
Thêm thuộc tính `Formula` vào database `07_Body_Weight_Tracker`:
```javascript
let(
  bmi, round((prop("Morning Weight (kg)") / ((175 / 100) * (175 / 100))) * 10) / 10,
  if(
    bmi < 18.5,
    "⚠️ BMI " + bmi + " (Thiếu cân - Cần thặng dư calo)",
    if(
      bmi <= 24.9,
      "✅ BMI " + bmi + " (Bình thường / Cân đối)",
      "⚡ BMI " + bmi + " (Thừa cân)"
    )
  )
)
```

### 3. Tự động tính Khối Lượng Tập Luyện (Volume Load)
Thêm thuộc tính `Formula` vào database `03_Gym_Workout_Logs`:
```javascript
if(
  empty(prop("Weight (kg)")) or empty(prop("Reps")),
  0,
  prop("Weight (kg)") * prop("Reps")
)
```

### 4. Tự động tính Calo món ăn theo Khối lượng nhập thực tế
Thêm thuộc tính `Formula` vào database `06_Nutrition_Daily_Meal_Log`:
```javascript
if(
  empty(prop("Amount")) or empty(prop("Calories (per 100g)")),
  prop("Calories (kcal)"),
  round((prop("Amount") * prop("Calories (per 100g)") / 100) * 10) / 10
)
```

---

## ⚡ HƯỚNG DẪN TẠO CÁC NÚT BẤM 1 CHẠM (NOTION BUTTONS)

Để bạn nhập liệu nhanh trên iPhone/Android mà không cần mở nhiều bảng:

1. **Nút `+ Cân Buổi Sáng`**:
   - Gõ `/button` $\rightarrow$ Đặt tên: `⚖️ + Cân Sáng Hôm Nay`.
   - Chọn hành động: `Add page to...` $\rightarrow$ Chọn database `07_Body_Weight_Tracker`.
   - Đặt giá trị mặc định: `Date` = `Today()`.
   - Thêm hành động tiếp theo: `Open page`. Khi bấm nút, Notion tự bật popup để bạn điền số kg vừa cân và giờ ngủ.

2. **Nút `+ Log Bữa Ăn`**:
   - Gõ `/button` $\rightarrow$ Đặt tên: `🥗 + Log Bữa Ăn`.
   - Chọn hành động: `Add page to...` $\rightarrow$ Chọn database `06_Nutrition_Daily_Meal_Log`.
   - Đặt giá trị: `Date` = `Today()`.
   - Thêm hành động: `Open page`.

3. **Nút `+ Ghi Hiệp Tập Hôm Nay`**:
   - Gõ `/button` $\rightarrow$ Đặt tên: `🏋️ + Ghi Hiệp Tập`.
   - Chọn hành động: `Add page to...` $\rightarrow$ Chọn database `03_Gym_Workout_Logs`.
   - Đặt giá trị: `Date` = `Today()`.

---

## 🎯 CHIẾN LƯỢC QUẢN LÝ DÀI HẠN
- **Không nhập thừa**: Chỉ cần log những ngày bạn thật sự ăn uống và tập luyện. Những ngày bận hoặc nghỉ ngơi chỉ cần ghi cân nặng buổi sáng.
- **Tập trung vào 2 con số cốt lõi**:
  1. Đường trung bình cân nặng tuần tăng **+0.2 đến +0.4 kg/tuần**.
  2. Mức tạ/reps của các bài chính (Incline DB Press, Lat Pulldown, Squat/Leg Press) nhích dần theo thời gian.
- **Đồng bộ hóa**: Bạn có thể tải app Notion trên điện thoại và đăng nhập cùng tài khoản để đồng bộ thời gian thực giữa máy tính và điện thoại.
