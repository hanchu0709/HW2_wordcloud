import tkinter as tk                        # 匯入標準 GUI 庫
from tkinter import filedialog, messagebox  # 匯入檔案選擇對話框與訊息提示框
from collections import Counter             # 匯入計數器，統計詞頻
from wordcloud import WordCloud             # 匯入文字雲生成庫
from PIL import Image, ImageTk              # 匯入圖像處理庫，將圖片轉成tkinter可相容的格式
import re                                   # 匯入正規表示式，清除不必要的文字字元
import numpy as np                          # 匯入數值計算庫，建立文字雲的圓形外框

# 定義停用詞集合
STOP_WORDS = {
    # 冠詞、介系詞、連字號 (基礎)
    "the", "a", "an", "and", "or", "but", "if", "because", "as", "until", "while",
    "of", "at", "by", "for", "with", "about", "against", "between", "into", "through",
    "during", "before", "after", "above", "below", "to", "from", "up", "down", "in",
    "out", "on", "off", "over", "under", "again", "further", "then", "once", "here", "there",

    # 人稱代名詞與所有格(I, you, your 等)
    "i", "me", "my", "myself", "we", "our", "ours", "ourselves", "you", "your", "yours", 
    "yourself", "yourselves", "he", "him", "his", "himself", "she", "her", "hers", 
    "herself", "it", "its", "itself", "they", "them", "their", "theirs", "themselves",

    # 助動詞與常見動詞
    "is", "are", "am", "was", "were", "be", "been", "being", "have", "has", "had", 
    "having", "do", "does", "did", "doing", "can", "could", "should", "will", "would", 
    "may", "might", "must", "get", "got", "go", "went", "say", "said", "make", "made",

    # 指示代名詞與疑問詞
    "this", "that", "these", "those", "who", "whom", "whose", "which", "what", 
    "when", "where", "why", "how", "all", "any", "both", "each", "few", "more", 
    "most", "other", "some", "such", "no", "nor", "not", "only", "own", "same", 
    "so", "than", "too", "very", "can", "just", "should", "now",

    # 常見口語詞與副詞
    "also", "really", "very", "well", "even", "still", "only", "maybe", "actually",
    "however", "therefore", "often", "always", "never", "something", "someone"

    # 增加縮寫後的殘留片段
    "dont", "didnt", "isnt", "arent", "wasnt", "werent", "hasnt", "havent", 
    "hadnt", "couldnt", "shouldnt", "wouldnt", "cant", "wont", "doesnt",
    "t", "s", "re", "ve", "ll", "d", "m", "don", "can" 
}

class WordCloudApp:
    def __init__(self, root):
        self.root = root
        self.root.title("進階英文文字雲產生器")
        self.root.geometry("750x900")
        self.result_win = None                                                                          # 初始化存放結果視窗的變數      

        # 主視窗大標題
        tk.Label(root, text="英文文字雲分析系統", font=("Microsoft JhengHei", 20, "bold")).pack(pady=15)

        # 設定要顯示的前n名字詞：top_n_entry
        top_n_frame = tk.Frame(root)                                                                        # 建立小框架(top_n_frame)用來水平排列標籤
        top_n_frame.pack(pady=5)                                                                            # 設定上下間距(pady)=5px 
        tk.Label(top_n_frame, text="設定分析高頻字詞：", font=("Microsoft JhengHei", 10)).pack(side=tk.LEFT)  # 在小框架中建立顯示「設定分析高頻字詞」的label；設置字體=10號微軟正黑體，並靠左排列(pack(side=tk.LEFT))
        self.top_n_entry = tk.Entry(top_n_frame, width=10)                                                  # 建立單行輸入框
        self.top_n_entry.insert(0, "20")                                                                    # 預設值=20
        self.top_n_entry.pack(side=tk.LEFT, padx=5)

        # 文字輸入區域：text_area     
        self.text_area = tk.Text(root, height=12, width=80, font=("Consolas", 10))  # 建立輸入方塊(Text)，設置高度=12、寬度=80、字體=10號Consolas
        self.text_area.pack(pady=10, padx=20)                                       # 設定輸入方塊與上方元件間距=10px、左右間距(padx)=20px

        # 存放按鈕的容器：btn_frame
        btn_frame = tk.Frame(root)  # 建立框架(Frame)作為容器，並水平排列其他功能按鈕
        btn_frame.pack(pady=10)     # 設定上下間距=10px
        
        # 建立功能按鈕：載入檔案load_file、產生文字雲create_result_window、清除並關閉clear_all
        tk.Button(btn_frame, text="載入檔案", command=self.load_file, width=12).pack(side=tk.LEFT, padx=10)                              # 建立「載入檔案」按鈕，點擊執行load_file，設置按鈕寬度12，並讓按鈕靠左排列
        tk.Button(btn_frame, text="產生文字雲", command=self.create_result_window, 
                  width=20, bg="#2196F3", fg="white", font=("Microsoft JhengHei", 10, "bold")).pack(side=tk.LEFT, padx=10)              # 建立「產生文字雲」按鈕，將按鈕設置為藍底白字，點擊執行create_result_window
        tk.Button(btn_frame, text="清除並關閉", command=self.clear_all, width=18, bg="#f44336", fg="white").pack(side=tk.LEFT, padx=10)  # 建立「清除並關閉」按鈕，將按鈕設置為紅底白字，點擊執行clear_all重置介面

        # 高頻字詞統計顯示區：stats_area
        tk.Label(root, text="【高頻字詞統計列表】", font=("Microsoft JhengHei", 12, "bold")).pack(pady=5)  # 建立「【高頻字詞統計列表】」的label作為標題，設置字體
        self.stats_area = tk.Text(root, height=15, width=60, bg="#f9f9f9", font=("Courier New", 10))     # 建立多行文字方塊(Text)顯示結果
        self.stats_area.pack(pady=10)

    # 讀取.txt檔案內容並顯示在輸入區
    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])  # 彈出檔案選擇對話框，限制只能選擇.txt的檔案
        if file_path:                                                                # file_path不為空
            try:
                with open(file_path, "r", encoding="utf-8") as file:                 # 以唯獨模式("r")開啟所選檔案，並指定utf-8編碼用以正確讀取文字
                    self.text_area.delete("1.0", tk.END)                             # 清空輸入區
                    self.text_area.insert(tk.END, file.read())                       # 插入.txt檔案文字內容到輸入區的末端
            except Exception as e:
                messagebox.showerror("錯誤", f"無法讀取檔案：{e}")

    # 文字前處理process_text：轉成小寫、過濾標點、去停用詞
    def process_text(self, text):
        text = text.lower()                                                            # 轉小寫
        text = text.replace("'", "")                                                   # 處理掉單引號，讓 don't 變成 dont
        clean_text = re.sub(r'[^a-z\s]', ' ', text)                                    # 將其他標點符號轉換為空格
        words = [w for w in clean_text.split() if w not in STOP_WORDS and len(w) > 1]  # 分詞與過濾
        return words

    def create_result_window(self):
        # 取得n值，並驗證其值是否為正整數
        try:
            top_n = int(self.top_n_entry.get().strip())                              # 從前n名的輸入框中取得文字，去除空白轉為整數
            if top_n <= 0: raise ValueError                                          # 若其值錯誤，引發ValueError
        except ValueError:                                                           # 執行ValueError
            messagebox.showwarning("輸入錯誤", "請在分析排名欄位輸入正整數（例如：20）")
            return

        # 檢查輸入區是否有文字input_text
        input_text = self.text_area.get("1.0", tk.END).strip()  # 取得輸入區全部內容並去除多餘換行與空白
        if not input_text:                                      # 若為空字串
            messagebox.showwarning("警告", "請先輸入內容")
            return

        words = self.process_text(input_text)                  # 呼叫process_text
        if not words:
            messagebox.showwarning("警告", "處理後沒有有效單詞")
            return

        # 進行字詞頻率統計並取出前n名
        word_counts = Counter(words)
        top_words_list = word_counts.most_common(top_n)  # 取得(單字,次數)的清單
        top_words_dict = dict(top_words_list)            # 轉為字典格式，供文字雲庫使用

        # 更新主介面的統計字詞頻率列表
        self.stats_area.delete("1.0", tk.END)
        header = f"{'排名':<5} {'單字 (Word)':<20} {'次數 (Freq)':<10}\n"  # 定義表格標題列，控制字串寬度(排名5格、單字20格、次數10格)，靠左對齊
        self.stats_area.insert(tk.END, header)                            # 從統計區插入定義好的標題列文字
        self.stats_area.insert(tk.END, "=" * 45 + "\n")
        for i, (word, count) in enumerate(top_words_list, 1):             # 遍歷字詞頻率清單，從1開始產生索引編號，i作為排名
            line = f"{i:<5} {word.ljust(20)} {count}\n"                   # 格式化每行內容：排名(i)寬度5、單字(word)、次數(count)
            self.stats_area.insert(tk.END, line)                          # 將格式化後的每行統計結果依序插入道統計區的末端

        # 製作圓形遮罩與文字雲圖片
        if self.result_win is not None and self.result_win.winfo_exists():
            self.result_win.destroy()                                      # 若舊式窗還在就先關閉

        # 建立1000*1000的矩陣，並繪製圓形邊界
        x, y = np.ogrid[:1000, :1000]                      # 將圓形以外的區域設為遮罩
        mask = (x - 500) ** 2 + (y - 500) ** 2 > 400 ** 2  # 轉換為文字雲庫可識別格式
        mask = 255 * mask.astype(int)

        try:
            # 根據統計出的前n名產生文字雲
            wc = WordCloud(
                width=1000, height=1000,                 # 設置原始寬度與高度           
                background_color="white",                # 背景顏色：白色
                mask=mask,                               # 使圖片變成圓形
                contour_width=3,                         # 設置外框線寬度
                contour_color='steelblue'                # 設置外框線顏色
            ).generate_from_frequencies(top_words_dict)  # 根據「前n名字詞頻率字典」計算文字大小與位置

            # 開啟新頂層視窗顯示文字雲圖片
            self.result_win = tk.Toplevel(self.root)             # 建立新的視窗(Toplevel)
            self.result_win.title(f"前 {top_n} 名高頻詞分析結果")  # 設置視窗標題
            self.result_win.geometry("850x920")                  # 設置視窗大小

            final_img = wc.to_image()                                                                 # 文字雲物件轉為image物件
            display_img = ImageTk.PhotoImage(final_img.resize((600, 600), Image.Resampling.LANCZOS))  # 調整圖片大小以符合視窗顯示

            img_label = tk.Label(self.result_win, image=display_img, bg="white")  # 建立label存放圖片
            img_label.image = display_img                                         # 手動保留圖片引用，避免被系統垃圾回收機制刪除導致圖片顯示空白
            img_label.pack(pady=10)                                               # 放置label並設置上下間距

            # 新視窗中的儲存按鈕
            tk.Button(self.result_win, text="儲存圖片", 
                      command=lambda: self.save_image(final_img),                              # 使用lambda函式傳遞當前圖片至save_image
                      bg="#4CAF50", fg="white", font=("Microsoft JhengHei", 12)).pack(pady=5)
        except Exception as e:
            messagebox.showerror("錯誤", f"生成錯誤：{e}")

    # 圖片儲存方法save_image
    def save_image(self, img_obj):
        path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG 檔案", "*.png")])  # 將圖片儲存至使用者選擇的路徑
        if path:
            img_obj.save(path)
            messagebox.showinfo("成功", "圖片已儲存")

    # 清空所有輸入框與關閉結果視窗
    def clear_all(self):
        self.text_area.delete("1.0", tk.END)
        self.stats_area.delete("1.0", tk.END)
        self.top_n_entry.delete(0, tk.END)
        self.top_n_entry.insert(0, "20")                                    # 重置為預設值
        if self.result_win is not None and self.result_win.winfo_exists():
            self.result_win.destroy()
            self.result_win = None

if __name__ == "__main__":
    root = tk.Tk()            # 建立根視窗
    app = WordCloudApp(root)  # 啟動應用程式類別
    root.mainloop()           # 進入視窗主循環