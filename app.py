import streamlit as st
import pandas as pd
import numpy as np

# ฟังก์ชันกรอง outlier ด้วย IQR
def filter_iqr(df, column):
    Q1 = df[column].quantile(0.25)
    Q3 = df[column].quantile(0.75)
    IQR = Q3 - Q1
    lower = Q1 - 1.5 * IQR
    upper = Q3 + 1.5 * IQR
    return df[(df[column] >= lower) & (df[column] <= upper)]

def main():
    st.title("วิเคราะห์เวลาล้างแม่พิมพ์ (Mold Washing Time)")

    uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (.xlsx)", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)

        # แปลงคอลัมน์วันที่-เวลาเป็น datetime
        for col in ["START WASHING DATE", "FINISH WASHING DATETIME", "TAKE IN DATETIME", "TAKE OUT DATETIME"]:
            df[col] = pd.to_datetime(df[col], errors='coerce')

        # คำนวณเวลาต่าง ๆ (หน่วยเป็นนาที)
        df["Time to wash (min)"] = (df["FINISH WASHING DATETIME"] - df["START WASHING DATE"]).dt.total_seconds() / 60
        df["Waiting Time In (min)"] = (df["START WASHING DATE"] - df["TAKE IN DATETIME"]).dt.total_seconds() / 60
        df["Waiting Time Out (min)"] = (df["TAKE OUT DATETIME"] - df["FINISH WASHING DATETIME"]).dt.total_seconds() / 60

        # กรองข้อมูล Status ที่ต้องการ
        valid_status = ["Send to production line", "Urgent", "Spear", "Return"]
        df = df[df["STATUS"].isin(valid_status)]

        # แยก PLANT ที่สนใจ (OS1, OS2-1, OS2-2)
        valid_plants = ["OS1", "OS2-1", "OS2-2"]
        df = df[df["PLANT"].isin(valid_plants)]

        # สร้างคอลัมน์วันที่จาก TAKE IN DATETIME เพื่อสรุปรายวัน
        df["Date"] = df["TAKE IN DATETIME"].dt.date

        # ฟังก์ชันกรอง IQR สำหรับแต่ละกลุ่ม (Plant, Status, Date)
        def apply_iqr_filter(group):
            for col in ["Time to wash (min)", "Waiting Time In (min)", "Waiting Time Out (min)"]:
                group = filter_iqr(group, col)
            return group

        # ใช้ apply เพื่อกรอง outlier ในแต่ละกลุ่ม
        df_filtered = df.groupby(["PLANT", "STATUS", "Date"]).apply(apply_iqr_filter).reset_index(drop=True)

        # สรุปค่าเฉลี่ยเวลาที่กรองแล้ว และนับจำนวน Mold จริง (นับ JOB NO ไม่ซ้ำ)
        summary = df_filtered.groupby(["PLANT", "STATUS", "Date"]).agg(
            avg_time_to_wash=pd.NamedAgg(column="Time to wash (min)", aggfunc="mean"),
            avg_waiting_in=pd.NamedAgg(column="Waiting Time In (min)", aggfunc="mean"),
            avg_waiting_out=pd.NamedAgg(column="Waiting Time Out (min)", aggfunc="mean"),
            mold_count=pd.NamedAgg(column="JOB NO", aggfunc=lambda x: x.nunique())
        ).reset_index()

        # แสดงผลตาราง
        st.subheader("สรุปเวลาต่อวัน (หลังกรอง outlier ด้วย IQR)")
        st.dataframe(summary.style.format({
            "avg_time_to_wash": "{:.2f} นาที",
            "avg_waiting_in": "{:.2f} นาที",
            "avg_waiting_out": "{:.2f} นาที",
            "mold_count": "{:d} ชุด"
        }))

        # กราฟตัวอย่าง: ค่าเฉลี่ย Time to wash แยกตาม Plant + Status
        import matplotlib.pyplot as plt
        import seaborn as sns

        st.subheader("กราฟเฉลี่ย Time to wash ต่อวัน")

        plant_filter = st.multiselect("เลือก Plant", options=valid_plants, default=valid_plants)
        status_filter = st.multiselect("เลือก Status", options=valid_status, default=valid_status)

        plot_data = summary[(summary["PLANT"].isin(plant_filter)) & (summary["STATUS"].isin(status_filter))]

        if not plot_data.empty:
            plt.figure(figsize=(12,6))
            sns.lineplot(data=plot_data, x="Date", y="avg_time_to_wash", hue="STATUS", style="PLANT", markers=True)
            plt.ylabel("เวลาเฉลี่ยล้าง (นาที)")
            plt.xticks(rotation=45)
            plt.tight_layout()
            st.pyplot(plt)
        else:
            st.write("ไม่มีข้อมูลที่ตรงกับตัวกรอง")

        # ปุ่มดาวน์โหลดสรุปเป็น CSV
        csv = summary.to_csv(index=False).encode('utf-8')
        st.download_button("ดาวน์โหลดผลสรุปเป็น CSV", csv, "summary.csv", "text/csv")

if __name__ == "__main__":
    main()
