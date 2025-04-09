import matplotlib.pyplot as plt

def plot_attendance(result_df):
    plt.figure(figsize=(10,6))
    result_df.plot(kind='bar', x='姓名', y='参与次数')
    plt.title("轮滑活动参与统计")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig("../docs/attendance_chart.png")