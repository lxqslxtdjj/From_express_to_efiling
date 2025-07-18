import re
import pandas as pd
from typing import List, Dict, Tuple
import logging
#step0
def part1():
    input_file = "PND53WHT.TXT"
    output_file = "step0-WTH.TXT"

    # 读取原始字节内容
    with open(input_file, "rb") as f:
        content = f.read()

    print("Before replacement (hex preview):")
    print(content[:100].hex())  # first 100 bytes in hex

    # ✅ 匹配 Line 008 类型的内容：SSA 开头，多个 VTS，ESA 结尾
    lines = content.split(b'\n')
    pattern_ssa = re.compile(b'(?:[\x00-\xff]*?\x8a){5,}')

    new_lines = []
    for line in lines:
        if pattern_ssa.search(line):
            new_lines.append(b'##########')  # 整行替换
        else:
            new_lines.append(line)

    content = b'\n'.join(new_lines)
    # ✅ 匹配 Line 44 类型的内容：BHP 开头，多个 HTS VTS，NBH 结尾
    lines = content.split(b'\n')
    pattern_bhp = re.compile(b'(?:[\x00-\xff]*?\x88){5,}')

    new_lines = []
    for line in lines:
        if pattern_bhp.search(line):
            new_lines.append(b'******')  # 整行替换
        else:
            new_lines.append(line)

    content = b'\n'.join(new_lines)

    # ✅ 替换所有 \x84 (IND) 为空格
    content = content.replace(b'\x84', b'\x20')
    
    print("After replacement (hex preview):")
    print(content[:300].hex())


    # 保存处理后的结果
    with open(output_file, "wb") as f:
        f.write(content)

    print(f"Cleaned file saved as {output_file}") 

#step1
def part2():
    def encode_1(input_file):
        # 尝试多种可能的编码方式
        encodings_to_try = ['tis-620', 'cp874','utf-8']
        
        for encoding in encodings_to_try:
            try:
                with open(input_file, 'rb') as f:
                    content = f.read().decode(encoding, errors='replace')
                    return content
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError("无法解码文件，尝试的编码都不适用")

    if __name__ == '__main__':
        input_file = 'step0-WTH.TXT'
        output_file = 'step1-WTH.TXT'
        # Step 1: Read and decode file
        content = encode_1(input_file)
        encode_1(input_file)
        
        # Step 2: Output the result to a new file
        with open("step1-WTH.TXT", "w", encoding="utf-8", newline='') as outfile:
            outfile.write(content)

        print(f"Processing complete. Output saved to '{output_file}'")

#step2
def part3():
    def extract_records(input_file, output_file):
        records = []
        in_record = False
        current_record = []

        # ✅ Read file content here
        with open(input_file, 'r', encoding='utf-8') as f:
            content = f.read()

        for line in content.split('\n'):
            # ✅ 如果你只想在记录中处理前导空格，就把这句移到下方 elif in_record 内
            # line = line.lstrip()  

            if '##########' in line:
                if in_record and current_record:  # 结束前一个记录
                    records.append('\n'.join(current_record)) 
                    current_record = []
                in_record = True
            elif '******' in line:
                in_record = False
                if current_record:  # 保存当前记录
                    records.append('\n'.join(current_record))
                    current_record = []
            elif in_record:
                current_record.append(line.lstrip())  # ✅ 只在记录中清理前导空格

        # 写入输出文件
        with open(output_file, "w", encoding="utf-8", newline='') as f:
            for record in records:
                f.write(record + '\n')
        
        print(f"成功提取 {len(records)} 条记录到 {output_file}")

    if __name__ == '__main__':
        input_file = 'step1-WTH.TXT'
        output_file = 'step2-WTH.TXT'
        extract_records(input_file, output_file)    
#step3
def part4():
    def main():
        # 示例文本内容（实际使用时替换为你的文件内容）
        with open('step2-WTH.TXT', 'r', encoding='utf-8') as f:
            content = f.read()
            content = content.replace('\ufffd', ' ')
        # 调用解析函数
        records = parse_records(content)
        # 去掉逗号
        cleaned_records = []
        for record in records:
            cleaned_record = {}
            for key, value in record.items():
                if isinstance(value, str):
                    cleaned_record[key] = value.replace(',', ' ').replace('，', ' ')
                else:
                    cleaned_record[key] = value  # 保留原值（数字等）
            cleaned_records.append(cleaned_record)

        # 转换为DataFrame
        df = pd.DataFrame(cleaned_records)
        
        # 保存为Excel
        df.to_excel('PND53output.xlsx', index=False)
        print("Excel文件已生成: PND53output.xlsx")

    def parse_records(content: str) -> List[Dict]:
        """解析文本记录并返回结构化数据"""
        if not content or not isinstance(content, str):
            raise ValueError("输入内容必须是非空字符串")

        # 预编译正则表达式（使用命名捕获组）
        pattern = re.compile(
            r'"\s*(?P<序号>\d+)\s+(?P<公司名称1>.*?)\s+'
            r'(?P<税号>\d+)\s+(?P<分支>\d+)\s+'
            r'(?P<日期>\d{2}/\d{2}/\d{2})\s+'
            r'(?P<服务类型>.*?)\s+'
            r'(?P<数量>\d+\.\d{2})\s+'
            r'(?P<未税金额>\d+\.\d{2})\s+'
            r'(?P<税额>\d+\.\d{2})\s+'
            r'(?P<报税类型>\d+)"\s*'
            r'(?P<公司名称2>.*?)\s*'
            r'"\s*(?P<地址1>.*?)\s*"\s*'
            r'(?P<地址2>.*?)\s*'
            r'(?="\s*\d+|$)',
            re.DOTALL | re.IGNORECASE
        )

        records = []
        
        for match in pattern.finditer(content):
            try:
                # 1. 处理税号（补全13位）
                tax_id = match.group('税号').zfill(13)
                
                # 2. 处理公司名称
                company_name = process_company_name(
                    match.group('公司名称1'), 
                    match.group('公司名称2')
                )
                
                # 3. 处理地址
                full_address, postal_code = process_address(
                    match.group('地址1'),
                    match.group('地址2')
                )
                
                # 4. 处理日期（泰历转换）
                thai_date = convert_to_thai_year(match.group('日期'))
                
                # 构建记录字典（按要求的顺序）
                record = {
                    'ลำดับที่': int(match.group('序号')),
                    'เลขประจำตัวผู้เสียภาษีอากร': tax_id,
                    'สาขาที่': int(match.group('分支')),
                    'คำนำหน้าชื่อ':' ',
                    'ชื่อ(100)': company_name,
                    **full_address,
                    'รหัสไปรษณีย์ (6)': postal_code,
                    'วันเดือนปีที่จ่าย': thai_date,
                    'ประเภทเงินได้ (200)': match.group('服务类型').strip(),
                    'อัตราภาษี (4 2)': float(match.group('数量')),
                    'จำนวนเงินได้ ( 15 2 )': float(match.group('未税金额')),
                    'จำนวนภาษีที่หัก ( 15 2 )': float(match.group('税额')),
                    'เงื่อนไข (1 )':int(match.group('报税类型'))
                }
                
                records.append(record)
                
            except (ValueError, AttributeError, TypeError) as e:
                logging.warning(f"处理记录时出错: {e}\n原始数据: {match.group()}")
                continue      
        return records
        
    def process_company_name(part1: str, part2: str) -> str:
        """处理公司名称：合并并标准化"""
        name = ' '.join(filter(None, [
            part1.strip(),
            part2.strip() if part2 else None
        ])).strip()
        
        # 替换公司类型缩写
        name = re.sub(
            r'\bco\.,?\s*ltd\b', 
            'company limited', 
            name, 
            flags=re.IGNORECASE
        )
        name = re.sub(
            r'\bco\.,?\s*l\.?t\.?d\.?\b', 
            'company limited', 
            name, 
            flags=re.IGNORECASE
        )
        
        return name

    def process_address(part1: str, part2: str) -> Tuple[dict, str]:
        """返回格式: ({ที่อยู่字段}, รหัสไปรษณีย์)"""
        # 合并地址
        full_address = ' '.join(filter(None, [
            part1.strip(),
            part2.strip() if part2 else None
        ])).strip()
        
        # 提取邮编
        postal_code = ''
        postal_match = re.search(r'(\d{5})(?!.*\d{5})', full_address)
        if postal_match:
            postal_code = postal_match.group(1)
            full_address= full_address[:postal_match.start()].strip()
        
        # 分割地址
        addr_dict = {
            'ที่อยู่ 1 (100)': full_address[:30].strip(),
            'ที่อยู่ 2 (100)': full_address[30:60].strip() if len(full_address) > 30 else '',
            'ที่อยู่ 3 (100)': full_address[60:90].strip() if len(full_address) > 60 else ''
        }
        
        return addr_dict, postal_code
        
    def convert_to_thai_year(date_str: str) -> str:
        """将短泰历年份转换为完整年份（BE佛历）"""
        try:
            month, day, year = date_str.split('/')
            # 假设68表示2568（BE佛历 = 西历 + 543）
            full_year = str(int(year) + 2500)  # 68 -> 2568
            return f"{month}/{day}/{full_year}"
        except:
            return date_str  # 如果转换失败返回原日期

    if __name__ == "__main__":
        # 配置日志记录
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        main()
#合并
if __name__ == "__main__":
    part1()
    part2()
    part3()
    part4()