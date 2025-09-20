import pandas as pd
import requests
import json
import time
import re

# --- 配置部分 ---
# Excel 文件路径
EXCEL_FILE_PATH = r"D:\work\跑图\015_开发认证集\开发验证集—综合-20250915.xlsx"
# Sheet 名称
SHEET_NAME = "prompt-v1"
# 包含要分类文本的列名（分词后的内容）
INPUT_COLUMN_NAME = "分词中文"
# 包含更完整描述的列名（问题/句子）
CONTEXT_COLUMN_NAME = "中文问题"
# 存放一级标签的列名
PRIMARY_LABEL_COLUMN = "一级标签"
# 存放二级标签的列名
SECONDARY_LABEL_COLUMN = "二级标签"
# API 端点 URL
API_URL = "http://10.155.111.41:28080/v1/chat/completions"
# 请求头
HEADERS = {
    "Content-Type": "application/json",
    "Authorization": "Bearer sk-cloud-media-123"
}
# 请求体模板
API_PAYLOAD_TEMPLATE = {
    "model": "Qwen/Qwen3-32B",
    "messages": [
        {"role": "user", "content": ""}
    ],
    "max_tokens": 1024,
    "temperature": 0.5
}
# 思考模式指令
THINKING_MODE_INSTRUCTION = "/nothink"
# 指令模板
INSTRUCTION_TEMPLATE = """
任务指令：标签匹配与分类
核心任务目标
你是一个专业的标签分类模型。你的核心任务是根据提供的完整句子和分词内容，在全面理解其语义的基础上，严格遵循预设的标签分类体系，为每个输入找到并输出唯一的一对一级标签和二级标签。

核心约束
唯一性：必须且只能输出一组一级标签和二级标签。

严格匹配：一级标签必须从以下六个中选择：实体、对象属性、全局属性、关系、场景、风格。

层级关系：二级标签必须严格从属于其对应的一级标签。

输入格式
你会接收到包含“完整句子”和“分词内容”的输入。
例如：

完整句子：橱柜是用暖色调的木头做的吗？

分词内容：橱柜、木材

这就需要你理解分词和句子最终的目的是要匹配什么标签给你，比如以上这个最有可能的匹配一级标签是对象属性，二级标签是对象属性下属的材质。橱柜、木材，不就说明他想要的是这个吗？

输出格式
输出格式必须严格遵循以下 JSON 模板，仅包含标签信息，不要包含任何其他说明文字或解释。

{
    "一级标签": "[一级分类名称]",
    "二级标签": "[二级分类名称]"
}


标签体系详情
一级分类：实体
二级标签：

人物（代称）：指代人物的称呼，如：妈、老婆、老头、美女。

人物（职业）：人物所从事的职业，如：学生、环卫工、医生、和尚。

动物

植物

静物：非生命体物品，如：石头、贝壳、化石、家电、家具、工具、玩具、摆件、雕像、文物。

普通建筑：常见的、非地标性的建筑，如：楼房、凉亭、教堂、高塔。

地标建筑：具有代表性或知名度的建筑，如：布达拉宫、比萨斜塔。

字符：文字、乐符等。

美食：可食用的食物和饮料，如：米饭、吐司、牛排、寿司、薯片、巧克力、水果、蔬菜。

器械：各种工具、机械和交通工具，如：船、火车、巴士、挖掘机、机械臂、枪支、火箭。

UI/UX：与用户界面/用户体验设计相关的内容。

现象：自然或非自然现象，如：火焰、闪电、彩虹、极光、涟漪。

一级分类：对象属性
二级标签：

相貌

人种

性别

年龄：包括确定年龄和年龄段，如：18岁、老年。

肤色

发型

体态：人或动物的体型或神态，如：苗条、臃肿、婀娜多姿。

装扮：人或动物的妆容、配饰、着装。

动作：人或动物的姿势、动作、活动等，如：站、挥舞、阅读。

情绪：人或动物的表情、情绪。

气质：人或动物所展现的个人特质，如：善良、可爱、庄严、高雅。

颜色：包括明确和泛意的颜色，如：红色、清新色彩。

数量：包括确数和概数，如：2个、一些。

纹饰：物体表面的纹理、图案、装饰。

样式：物体的外观形态，如：形状、大小、长短。

材质：棉、麻、塑料、金属等。

状态：旧、坏、脏等非活动性状态。

动词：描述动作的词语，如：旋转、流动、滑行。

一级分类：全局属性
二级标签：

时代：图片的时代背景，如：中世纪、90年代、二战时期。

氛围：通过视听触所感受到的环境气氛，如：冷清、阴森。

意境：人对环境的纯主观感受，如：神秘、华丽。

视角：图片的拍摄或描绘角度，如：前视、航拍、仰视、居中。

景别：描述画面距离，如：近景、远景、特写、半身、全身照。

镜头：描述镜头角度，如：鱼眼镜头、微距、超广角。

光影

调性：图片的整体风格倾向，如：小清新、厚重、年代感。

精度：表示图片质量，如：高清、HDR、高分辨率。

虚拟：抽象概念、虚构的事物、电脑网络。

神话/童话：内容所涉及的宏观主题，如：神话、童话、科幻。

一级分类：关系
二级标签：

位置关系：描述物体间的相对位置，如：前后左右、层叠、交叉。

交互关系：描述物体间的互动，如：手拿东西、拥抱、握手、吹动、吸引。

一级分类：场景
二级标签：

室内环境：室内地点，如：厨房、卧室、展厅、影院、超市。

自然环境：自然界地点，如：海边、森林、雪山、戈壁、草原。

人文环境：人类活动地点，如：城市街道、村庄、广场、足球场、泳池。

生活场景：描绘生活事件，如：聚餐、毕业典礼、婚礼、派对、比赛。

自然场景：描绘自然事件，如：动物大迁徙、暴风雨。

一级分类：风格
二级标签：

绘画

摄影

漫画

插画

单反
"""

def get_labels_from_api(full_text, classified_text):
    """
    通过 API 获取一级和二级标签。
    
    Args:
        full_text (str): 完整的句子或问题。
        classified_text (str): 需要分类的分词文本。
    
    Returns:
        tuple: (一级标签, 二级标签) 或 (None, None) 如果请求失败。
    """
    # 结合指令和要分类的文本
    full_prompt = f"{THINKING_MODE_INSTRUCTION}\n{INSTRUCTION_TEMPLATE}\n完整句子：{full_text}\n分词内容：{classified_text}"
    API_PAYLOAD_TEMPLATE["messages"][0]["content"] = full_prompt
    
    try:
        response = requests.post(API_URL, headers=HEADERS, json=API_PAYLOAD_TEMPLATE, timeout=60)
        response.raise_for_status()  # 如果状态码不是 200，则抛出异常
        
        result = response.json()
        raw_content = result['choices'][0]['message']['content']
        
        # 尝试从返回的字符串中提取 JSON
        match = re.search(r"\{.*?\}", raw_content, re.DOTALL)
        if match:
            json_str = match.group(0)
            data = json.loads(json_str)
            
            primary_label = data.get("一级标签", "解析失败")
            secondary_label = data.get("二级标签", "解析失败")
            
            return primary_label, secondary_label
        else:
            print(f"API 返回内容中未找到有效的 JSON：\n{raw_content}")
            return "解析失败", "解析失败"

    except requests.exceptions.RequestException as e:
        print(f"请求失败: {e}")
        return "请求失败", "请求失败"
    except json.JSONDecodeError as e:
        print(f"JSON 解析失败: {e}\n原始内容:\n{raw_content}")
        return "JSON解析错误", "JSON解析错误"
    except Exception as e:
        print(f"处理结果时发生错误: {e}")
        return "未知错误", "未知错误"

def process_excel_file():
    """
    读取 Excel 文件，调用 API，并每处理 100 行将结果保存一次。
    """
    try:
        # 读取 Excel 文件
        df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=SHEET_NAME)
        print(f"成功读取文件: {EXCEL_FILE_PATH}, Sheet: {SHEET_NAME}")

        # 检查必要的列是否存在
        required_columns = [INPUT_COLUMN_NAME, CONTEXT_COLUMN_NAME, PRIMARY_LABEL_COLUMN, SECONDARY_LABEL_COLUMN]
        if not all(col in df.columns for col in required_columns):
            print(f"错误：Excel 文件中缺少必要的列。请确保包含 {required_columns}。")
            return
            
        # --- 关键修改：添加计数器 ---
        save_interval = 100  # 每处理 100 行保存一次
        rows_processed_since_last_save = 0
        # --- 计数器修改结束 ---

        for index, row in df.iterrows():
            text_to_classify = row[INPUT_COLUMN_NAME]
            full_context = row[CONTEXT_COLUMN_NAME]
            
            # 如果输入内容或上下文内容为空，则跳过
            if (not isinstance(text_to_classify, str) or pd.isna(text_to_classify)) or \
               (not isinstance(full_context, str) or pd.isna(full_context)):
                print(f"第 {index+2} 行 '{INPUT_COLUMN_NAME}' 或 '{CONTEXT_COLUMN_NAME}' 列为空，跳过。")
                continue
            
            print(f"正在处理第 {index+2} 行：\n完整句子: '{full_context}'\n分词内容: '{text_to_classify}'")
            
            # 调用 API 获取标签
            primary_label, secondary_label = get_labels_from_api(full_context, text_to_classify)
            
            # 将结果更新到 DataFrame
            df.at[index, PRIMARY_LABEL_COLUMN] = primary_label
            df.at[index, SECONDARY_LABEL_COLUMN] = secondary_label
            
            print(f"已更新：一级标签 -> {primary_label}, 二级标签 -> {secondary_label}")
            
            # 增加计数器
            rows_processed_since_last_save += 1
            
            # --- 关键修改：检查是否达到保存间隔 ---
            if rows_processed_since_last_save >= save_interval:
                try:
                    df.to_excel(EXCEL_FILE_PATH, sheet_name=SHEET_NAME, index=False)
                    print(f"已处理 {rows_processed_since_last_save} 行，文件已保存。")
                    rows_processed_since_last_save = 0  # 重置计数器
                except Exception as e:
                    print(f"保存文件失败: {e}。请确保文件未被其他程序占用。")
            # --- 间隔保存修改结束 ---
            
            # 简单延迟，避免对 API 造成过大压力
            time.sleep(1) 

        # --- 关键修改：循环结束后保存剩余的数据 ---
        if rows_processed_since_last_save > 0:
            try:
                df.to_excel(EXCEL_FILE_PATH, sheet_name=SHEET_NAME, index=False)
                print(f"\n所有行处理完成，剩余 {rows_processed_since_last_save} 行已保存。")
            except Exception as e:
                print(f"保存文件失败: {e}。请确保文件未被其他程序占用。")
        else:
            print("\n所有行处理完成，文件已保存。")
        # --- 结束保存修改结束 ---
        
    except FileNotFoundError:
        print(f"错误：文件未找到，请检查路径: {EXCEL_FILE_PATH}")
    except Exception as e:
        print(f"发生未知错误: {e}")

if __name__ == "__main__":
    process_excel_file()