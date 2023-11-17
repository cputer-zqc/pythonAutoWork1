import tkinter as tk
from tkinter import ttk

class PartitionedWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("上下排列的分区示例")

        # 创建第一个分区
        partition1 = ttk.Labelframe(self, text="分区1")
        label1 = tk.Label(partition1, text="这是第一个分区的内容")
        label1.pack(padx=10, pady=10)
        partition1.pack(side=tk.TOP, padx=10, pady=10, fill=tk.BOTH, expand=True)

        # 创建第二个分区
        partition2 = ttk.Labelframe(self, text="分区2")
        label2 = tk.Label(partition2, text="这是第二个分区的内容")
        label2.pack(padx=10, pady=10)
        partition2.pack(side=tk.TOP, padx=10, pady=10, fill=tk.BOTH, expand=True)

        # 创建第三个分区
        partition3 = ttk.Labelframe(self, text="分区3")
        label3 = tk.Label(partition3, text="这是第三个分区的内容")
        label3.pack(padx=10, pady=10)
        partition3.pack(side=tk.TOP, padx=10, pady=10, fill=tk.BOTH, expand=True)

if __name__ == "__main__":
    app = PartitionedWindow()
    app.mainloop()
