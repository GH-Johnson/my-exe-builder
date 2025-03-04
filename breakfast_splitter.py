import pandas as pd
import os
from tkinter import Tk, filedialog, messagebox

def select_file():
    """ 选择原始Excel文件 """
    root = Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="选择早餐名单文件",
        filetypes=[("Excel文件", "*.xlsx")]
    )

def select_output_dir():
    """ 选择输出目录 """
    root = Tk()
    root.withdraw()
    return filedialog.askdirectory(
        title="选择结果保存位置"
    )

def split_excel():
    try:
        # 文件选择
        input_file = select_file()
        if not input_file: return
        output_dir = select_output_dir()
        if not output_dir: return

        # 读取数据
        df = pd.read_excel(input_file, sheet_name='订单明细')
        df = df[df['所属组'] == '学生组']
        df = df[['用户名称', '所属部门', '下单时间']]
        
        # 数据预处理
        df['年级'] = df['所属部门'].str[0]
        df['班级'] = df['所属部门'].str[1:].str.replace('/', '_')  # 清理班级名称
        df['订单日期'] = pd.to_datetime(df['下单时间']).dt.strftime('%Y%m%d')  # 新增日期列
        
        # 创建多级分组
        grouped = df.groupby(['年级', '订单日期', '班级'])
        
        # 生成文件
        os.makedirs(output_dir, exist_ok=True)
        file_counter = 0
        
        for (grade, order_date, class_name), group in grouped:
            # 构建文件名
            filename = f"{grade}年级_{order_date}.xlsx"
            file_path = os.path.join(output_dir, filename)
            
            # 写入模式判断
            mode = 'a' if os.path.exists(file_path) else 'w'
            
            with pd.ExcelWriter(file_path, engine='openpyxl', mode=mode) as writer:
                group[['用户名称', '下单时间']].to_excel(
                    writer, 
                    sheet_name=class_name[:31],  # 遵守Excel 31字符限制
                    index=False
                )
                file_counter += 1

        messagebox.showinfo("完成", 
            f"生成{file_counter}个带日期文件\n保存至：{output_dir}")
    
    except Exception as e:
        messagebox.showerror("错误", f"处理失败：{str(e)}")

if __name__ == '__main__':
    split_excel()