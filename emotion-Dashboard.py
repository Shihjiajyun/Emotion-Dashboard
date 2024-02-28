import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge

def replace_emotions(emotion):
    if emotion == "積極心情":
        return 2
    elif emotion == "中性心情":
        return 1
    elif emotion == "複雜心情":
        return -1
    elif emotion == "消極心情":
        return -2
    else:
        return emotion

def multiply_values(row):
    try:
        values = [float(value) for value in row] 
        result = 1
        for value in values:
            result *= value
        return result
    except ValueError:
        return None

def read_excel(file_path):
    try:
        df = pd.read_excel(file_path, usecols=[1, 2, 3])
        df.columns = ['Emotion', 'Intensity (1-10)', 'Duration (hours)']
        df['Emotion'] = df['Emotion'].apply(replace_emotions)
        df['Emotion Score'] = df.apply(multiply_values, axis=1)
        df = df.dropna()  # 刪除NaN值
        total_score = df['Emotion Score'].sum()

        # 打印出数据
        print("數據：\n", df)

        # 绘制仪表板
        plot_gauge(total_score)

    except FileNotFoundError:
        print("找不到指定的文件。")

def plot_gauge(total_score):
    min_value = 0
    max_value = 1000
    angle = 180 - (180 / max_value) * total_score  

    if 0 <= total_score < 250:
        emotion_label = "Negative Emotion"
    elif 250 <= total_score < 500:
        emotion_label = "Complex Emotion"
    elif 500 <= total_score < 750:
        emotion_label = "Neutral Emotion"
    else:
        emotion_label = "Positive Emotion"

    fig, ax = plt.subplots(figsize=(8, 8))
    ax.set_xlim(-1.5, 1.5)
    ax.set_ylim(0, 1.5)
    ax.axis('off')
    ax.set_aspect('equal')


    circle = Wedge((0, 0), 1, 0, 180, color='lightgrey', ec='black', lw=2) 
    ax.add_patch(circle)

    pointer = plt.Line2D((0, 0.5 * np.cos(np.radians(angle))),
                         (0, 0.5 * np.sin(np.radians(angle))),
                         lw=2, color='red')
    ax.add_line(pointer)


    ax.text(-1.5, -0.1, str(min_value), fontsize=12, ha='center')
    ax.text(1.5, -0.1, str(max_value), fontsize=12, ha='center')
    ax.text(0, -0.3, f'Total Score: {total_score}', fontsize=14, ha='center')
    ax.text(0, -0.5, f'Emotion: {emotion_label}', fontsize=14, ha='center', color='blue')

    colors = ['#E98A15', '#ECE5F0', '#003B36', '#012622']
    angles = [0, 45, 90, 135, 180]
    labels = ['Positive', 'Neutral', 'Complex', 'Negative']
    for i in range(len(angles) - 1):
        wedge = Wedge((0, 0), 1, angles[i], angles[i + 1], color=colors[i], ec='black', lw=2)
        ax.add_patch(wedge)
        ax.text(1.2 * np.cos(np.radians((angles[i] + angles[i + 1]) / 2)),
                1.2 * np.sin(np.radians((angles[i] + angles[i + 1]) / 2)),
                labels[i], fontsize=10, ha='center', va='center', color='black')

    plt.title('Emotion Dashboard')
    plt.show()

def main():
    file_path = input("輸入文件路徑：")
    read_excel(file_path)

if __name__ == "__main__":
    main()
