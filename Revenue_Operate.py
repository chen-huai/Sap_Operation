import datetime
from chinese_calendar import is_holiday
import pandas as pd
import os

# 新增计算方法
class RevenueAllocator:
    def __init__(self):
        self.hours_data = pd.DataFrame(columns=['date', 'staff_name', 'hours', 'order_no', 'dept'])
        self.hours_file = None  # 将在使用时根据日期确定文件名

    def _get_hours_file_path(self, date, path):
        """
        根据日期生成工时数据文件路径
        :param date: datetime.date对象
        :return: 文件路径
        """
        month_str = date.strftime("%Y%m")
        return f"{path}\\hours_{month_str}.csv"

    def _load_hours_data(self, date, path):
        """
        加载指定月份的工时数据
        :param date: datetime.date对象，用于确定月份
        """
        self.hours_file = self._get_hours_file_path(date, path)
        if os.path.exists(self.hours_file):
            try:
                self.hours_data = pd.read_csv(self.hours_file)
                # 确保日期列是datetime类型
                self.hours_data['date'] = pd.to_datetime(self.hours_data['date']).dt.date
            except Exception as e:
                print(f"Error loading hours data: {e}")
                self.hours_data = pd.DataFrame(columns=['date', 'staff_name', 'hours', 'order_no', 'dept'])

    def _save_hours_data(self):
        """
        保存工时数据到对应月份的文件
        """
        if self.hours_file is None:
            return
            
        try:
            self.hours_data.to_csv(self.hours_file, index=False)
        except Exception as e:
            print(f"Error saving hours data: {e}")

    def _get_staff_daily_hours(self, date, staff_name):
        """
        获取指定员工在指定日期的工作时长
        """
        if self.hours_data.empty:
            return 0
        
        mask = (self.hours_data['date'] == date) & (self.hours_data['staff_name'] == staff_name)
        return self.hours_data.loc[mask, 'hours'].sum()

    def _update_staff_daily_hours(self, date, staff_name, hours, order_no, dept):
        """
        更新指定员工在指定日期的工作时长
        """
        # 添加新的工时记录
        new_record = pd.DataFrame({
            'date': [date],
            'staff_name': [staff_name],
            'hours': [hours],
            'order_no': [order_no],
            'dept': [dept]
        })
        
        # 合并到现有数据
        self.hours_data = pd.concat([self.hours_data, new_record], ignore_index=True)
        
        # 保存更新后的数据
        self._save_hours_data()

    def _get_available_hours(self, date, staff_name, max_hours_per_day):
        """
        获取指定员工在指定日期可用的工作时长
        """
        current_hours = self._get_staff_daily_hours(date, staff_name)
        return max(0, max_hours_per_day - current_hours)

    def calculate_revenue_allocation(self, revenueData, configContent):
        """动态配置的收入分配计算方法"""
        material_code = revenueData.get('materialCode', '')
        base = (float(revenueData['Revenue']) * float(configContent.get('Plan_Cost_Parameter')) - float(revenueData['Total Subcon Cost']) / 1.06)

        result = {
            'business_dept_1000_revenue': 0, 'lab_1000_revenue': 0, 'business_dept_1000_hours': 0, 'lab_1000_hours': 0,
            'business_dept_2000_revenue': 0, 'lab_2000_revenue': 0, 'business_dept_2000_hours': 0, 'lab_2000_hours': 0,
            'lab_1000': '', 'lab_2000': '', 'order_no': revenueData.get('orderNo', ''),
            'material_code_1000': '',
            'material_code_2000': ''
        }

        # 统一使用 Business_Department 配置
        business_dept = configContent.get('Business_Department', 'CS')

        # 情况1: 配置不存在
        if material_code not in configContent:
            prefix = material_code.split('-')[0]
            lab = configContent.get(prefix)

            lab_cost = float(configContent.get(f"{lab}_Cost_Parameter", 0.3))
            lab_rate = float(configContent.get(f"{lab}_Hourly_Rate", 342))
            business_dept_rate = float(configContent.get(f"{business_dept}_Hourly_Rate", 315))

            result.update({
                'business_dept_1000_revenue': base * (1 - lab_cost),
                'lab_1000_revenue': base * lab_cost,
                'business_dept_1000_hours': (base * (1 - lab_cost)) / business_dept_rate,
                'lab_1000_hours': (base * lab_cost) / lab_rate,
                'lab_1000': lab,
                'material_code_1000': material_code
            })
        else:
            # 情况2: 存在特殊配置
            rule = configContent.get(material_code, 'PHY_1000/CHM_2000').split('/')
            item_type = material_code.split('-')[1]

            proportion_1000 = float(configContent.get(f"{item_type}_Item_1000", 0.8))
            proportion_2000 = float(configContent.get(f"{item_type}_Item_2000", 0.2))

            lab_1000 = rule[0].split('_')[0]
            lab_2000 = rule[1].split('_')[0]

            # 获取实验室参数
            lab_1000_cost = float(configContent.get(f"{lab_1000}_Cost_Parameter", 0.3))
            lab_2000_cost = float(configContent.get(f"{lab_2000}_Cost_Parameter", 0.3))
            lab_1000_rate = float(configContent.get(f"{lab_1000}_Hourly_Rate", 342))
            lab_2000_rate = float(configContent.get(f"{lab_2000}_Hourly_Rate", 342))
            business_dept_rate = float(configContent.get(f"{business_dept}_Hourly_Rate", 315))

            # 计算双部门分配
            result.update({
                'business_dept_1000_revenue': base * proportion_1000 * (1 - lab_1000_cost),
                'lab_1000_revenue': base * proportion_1000 * lab_1000_cost,
                'business_dept_2000_revenue': base * proportion_2000 * (1 - lab_2000_cost),
                'lab_2000_revenue': base * proportion_2000 * lab_2000_cost,
                'business_dept_1000_hours': (base * proportion_1000 * (1 - lab_1000_cost)) / business_dept_rate,
                'lab_1000_hours': (base * proportion_1000 * lab_1000_cost) / lab_1000_rate,
                'business_dept_2000_hours': (base * proportion_2000 * (1 - lab_2000_cost)) / business_dept_rate,
                'lab_2000_hours': (base * proportion_2000 * lab_2000_cost) / lab_2000_rate,
                'lab_1000': lab_1000,
                'lab_2000': lab_2000,
                'material_code_1000': configContent.get(f"{material_code}_mc").split('/')[0],
                'material_code_2000': configContent.get(f"{material_code}_mc").split('/')[1]
            })

        # 在返回结果前增加数据结构整理
        return [
            {  # 1000业务部门
                'order_no': revenueData['orderNo'],
                'material_code': result.get("material_code_1000"),  # 添加后缀
                'item': '1000',
                'dept': business_dept,
                'dept_revenue': result['business_dept_1000_revenue'],
                'dept_hours': result['business_dept_1000_hours'],
                'original_hours': result['business_dept_1000_hours']  # 保存原始工时
            },
            {  # 1000实验室
                'order_no': revenueData['orderNo'],
                'material_code': result.get("material_code_1000"),  # 相同分组共用后缀
                'item': '1000',
                'dept': result['lab_1000'],
                'dept_revenue': result['lab_1000_revenue'],
                'dept_hours': result['lab_1000_hours'],
                'original_hours': result['lab_1000_hours']  # 保存原始工时
            },
            {  # 2000业务部门
                'order_no': revenueData['orderNo'],
                'material_code': result.get("material_code_2000"),  # 添加不同后缀
                'item': '2000',
                'dept': business_dept,
                'dept_revenue': result['business_dept_2000_revenue'],
                'dept_hours': result['business_dept_2000_hours'],
                'original_hours': result['business_dept_2000_hours']  # 保存原始工时
            },
            {  # 2000实验室
                'order_no': revenueData['orderNo'],
                'material_code': result.get("material_code_2000"),  # 相同分组共用后缀
                'item': '2000',
                'dept': result['lab_2000'],
                'dept_revenue': result['lab_2000_revenue'],
                'dept_hours': result['lab_2000_hours'],
                'original_hours': result['lab_2000_hours']  # 保存原始工时
            }
        ]

    # 新增工作日生成方法
    def generate_work_days(self, start_date, end_date):
        """生成有效工作日列表（自动排除节假日和周末）"""
        from chinese_calendar import is_holiday
        work_days = []
        current_day = start_date
        while current_day <= end_date:
            if not is_holiday(current_day) and current_day.weekday() < 5:
                work_days.append(current_day)
            current_day += datetime.timedelta(days=1)
        return work_days

    def _get_week_number(self, date):
        """
        获取日期在当年的周数
        :param date: datetime.date对象
        :return: 周数 (int)
        """
        return date.isocalendar()[1]

    def allocate_working_hours(self, results, max_hours_per_day, start_date, end_date, staff_dict, configContent):
        """
        增强版工时分配方法（集成节假日和人员分配）
        :param results: 需要分配工时的记录列表
        :param max_hours_per_day: 每人每天最大工作时长
        :param start_date: 开始日期
        :param end_date: 结束日期
        :param staff_dict: 部门人员字典 {部门: [员工编号1, 员工编号2...]}
        :return: 分配后的工时记录列表
        """
        # 加载现有工时数据（使用开始日期确定月份）
        self._load_hours_data(start_date, configContent.get('Hour_Files_Export_URL'))

        # 过滤零工时数据
        filtered = [r for r in results if r['dept_hours'] > 0]

        # 生成工作日历
        work_days = self.generate_work_days(start_date, end_date)
        if not work_days:
            return []

        # 初始化部门人员轮询指针
        dept_pointers = {dept: 0 for dept in staff_dict.keys()}
        # 记录每个订单当前分配的员工
        order_staff = {}  # {order_no: staff_name}

        final_results = []
        
        # 按订单号和部门分组处理工时
        order_dept_groups = {}
        for record in filtered:
            key = (record['order_no'], record['dept'])
            if key not in order_dept_groups:
                order_dept_groups[key] = []
            order_dept_groups[key].append(record)

        # 处理每个订单组
        for (order_no, dept), order_records in order_dept_groups.items():
            # 获取对应部门的员工列表
            staff_list = staff_dict.get(dept, [])
            
            if not staff_list:
                print(f"Warning: No staff found for department {dept}")
                continue

            # 获取当前订单的已分配员工（如果有）
            current_staff_name = order_staff.get(order_no)
            
            # 遍历每个工作日进行分配
            for work_day in work_days:
                # 如果当前订单已经有分配的员工，优先使用该员工
                if current_staff_name and current_staff_name in staff_list:
                    available_hours = self._get_available_hours(work_day, current_staff_name, max_hours_per_day)
                    if available_hours > 0:
                        # 尝试分配最大可能的工时
                        allocate_hours = min(available_hours, sum(record['dept_hours'] for record in order_records))
                        if allocate_hours > 0:
                            # 分配工时给当前员工
                            for record in order_records:
                                if record['dept_hours'] <= 0:
                                    continue
                                
                                # 计算该记录可分配的工时
                                record_hours = min(allocate_hours, record['dept_hours'])
                                if record_hours <= 0:
                                    continue
                                
                                new_record = record.copy()
                                new_record.update({
                                    'allocated_date': work_day,
                                    'allocated_hours': record_hours,
                                    'staff_name': current_staff_name,
                                    'staff_id': configContent.get(current_staff_name),
                                    'week': self._get_week_number(work_day)
                                })
                                final_results.append(new_record)
                                
                                # 更新工时记录
                                self._update_staff_daily_hours(work_day, current_staff_name, record_hours, order_no, dept)
                                record['dept_hours'] -= record_hours
                                allocate_hours -= record_hours
                                
                                if allocate_hours <= 0:
                                    break
                        continue
                
                # 如果没有已分配的员工或当前员工已满，寻找新员工
                staff_namex = dept_pointers[dept] % len(staff_list)
                for _ in range(len(staff_list)):  # 尝试所有员工
                    staff_name = staff_list[staff_namex]
                    available_hours = self._get_available_hours(work_day, staff_name, max_hours_per_day)
                    
                    if available_hours > 0:
                        # 尝试分配最大可能的工时
                        allocate_hours = min(available_hours, sum(record['dept_hours'] for record in order_records))
                        if allocate_hours > 0:
                            # 分配工时给新员工
                            for record in order_records:
                                if record['dept_hours'] <= 0:
                                    continue
                                
                                # 计算该记录可分配的工时
                                record_hours = min(allocate_hours, record['dept_hours'])
                                if record_hours <= 0:
                                    continue
                                
                                new_record = record.copy()
                                new_record.update({
                                    'allocated_date': work_day,
                                    'allocated_hours': record_hours,
                                    'staff_name': staff_name,
                                    'staff_id': configContent.get(staff_name),
                                    'week': self._get_week_number(work_day)
                                })
                                final_results.append(new_record)
                                
                                # 更新工时记录
                                self._update_staff_daily_hours(work_day, staff_name, record_hours, order_no, dept)
                                record['dept_hours'] -= record_hours
                                allocate_hours -= record_hours
                                
                                # 记录该订单的分配员工
                                order_staff[order_no] = staff_name
                                
                                if allocate_hours <= 0:
                                    break
                            break
                    
                    staff_namex = (staff_namex + 1) % len(staff_list)
                
                dept_pointers[dept] += 1

        return final_results


# if __name__ == "__main__":
#     # 示例数据
#
#     revenueDatas = [
#         {'orderNo': 'ORD123',
#          'materialCode': 'T20-430-A2',
#          'Revenue': 20000,
#          'Total Subcon Cost': 2000},
#         {'orderNo': 'ORD124',
#          'materialCode': 'T75-405-A2',
#          'Revenue': 10000,
#          'Total Subcon Cost': 2000},
#         {'orderNo': 'ORD125',
#          'materialCode': 'T75-441-A2',
#          'Revenue': 30000,
#          'Total Subcon Cost': 2000},
#         {'orderNo': 'ORD126',
#          'materialCode': 'T20-441-00',
#          'Revenue': 35000,
#          'Total Subcon Cost': 2000},
#
#     ]
#     configContent = {'特殊开票': '内容',
#                      'SAP_Date_URL': 'N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\收样\\3.Sap\\ODM Data - XM',
#                      'Invoice_File_URL': 'N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\收样\\3.Sap\\ODM Data - XM\\2.特殊开票',
#                      'Invoice_File_Name': '特殊开票要求2022.xlsx', 'Data数据处理': '内容',
#                      'Row Data': 'Client Contact Name',
#                      'Column Data': 'Project No.;Currency;Amount with VAT;Reference No.', 'Row Check': '0',
#                      'Column Check': '0',
#                      'Combine Key': "CS;Sales;Currency;Material Code;Invoices' name (Chinese);Buyer(GPC);Month;Exchange Rate",
#                      'SAP登入信息': '内容', 'Login_msg': 'DR-0486-01->601-240', 'Business_Department': 'CS',
#                      'Lab_1': 'PHY', 'Lab_2': 'CHM', 'T20': 'PHY', 'T75': 'CHM', 'Hourly Rate': '金额',
#                      'CS_Hourly_Rate': '300', 'PHY_Hourly_Rate': '300', 'CHM_Hourly_Rate': '300', '成本中心': '编号',
#                      'CS_Selected': '1', 'PHY_Selected': '1', 'CHM_Selected': '1', 'CS_Cost_Center': '48601240',
#                      'CHM_Cost_Center': '48601293', 'PHY_Cost_Center': '48601294', '计划成本': '数值',
#                      'Plan_Cost_Parameter': '0.9', 'Significant_Digits': '0', '实验室成本比例': '数值',
#                      'CHM_Cost_Parameter': '0.3', 'PHY_Cost_Parameter': '0.3', '405_Item_1000': '0.5',
#                      '405_Item_2000': '0.5', '441_Item_1000': '0.8', '441_Item_2000': '0.2', '430_Item_1000': '0.8',
#                      '430_Item_2000': '0.2', 'T20-430-A2': 'PHY_1000/CHM_2000',
#                      'T20-430-A2_mc': 'T20-430-00/T75-430-00', 'T75-441-A2': 'CHM_1000/PHY_2000',
#                      'T75-441-A2_mc': 'T75-441-00/T20-441-00', 'T75-405-A2': 'CHM_1000/PHY_2000',
#                      'T75-405-A2_mc': 'T75-405-00/T20-405-00', 'Max_Hour': '8', 'DATA A数据填写': '判断依据',
#                      'Data_A_E1': '5010815347;5010427355;5010913488;5010685589;5010829635;5010817524',
#                      'Data_A_Z2': '5010908478;5010823259', 'SAP操作': '内容', 'Cost_VAT_Selected': '1',
#                      'NVA01_Selected': '1', 'NVA02_Selected': '1', 'NVF01_Selected': '0', 'NVF03_Selected': '0',
#                      'DataB_Selected': '1', 'Plan_Cost_Selected': '25', 'Save_Selected': '1', 'Every_Selected': '1',
#                      'Contact_Selected': '0', '管理操作': '内容',
#                      'Billing_List_URL': 'N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\2.财务\\3.Billing存档\\4.XM-billing list\\2023',
#                      'Add_CS_Msg_Selected': '1', 'Invoice_No_Selected': '1', 'Invoice_Start_Num': '4',
#                      'Invoice_Num': '9', 'Company_Name_Selected': '1', 'Order_No_Selected': '0',
#                      'Invoice_Contact_Selected': '0', 'Order_Start_Num': '7', 'Order_Num': '9',
#                      'Project_No_Selected': '0', 'Invoice_Name': 'CS + Invoice No + Company Name',
#                      'Invoice_Files_Import_URL': 'C:\\Users\\chen-fr\\Desktop',
#                      'Invoice_Files_Export_URL': 'N:\\XM Softlines\\1. Project\\3. Finance\\02. WIP',
#                      'Ele_Invoice_No_Selected': '1', 'Ele_Invoice_Start_Num': '486', 'Ele_Invoice_Num': '9',
#                      'Ele_Order_No_Selected': '0', 'Ele_Order_Start_Num': '7486', 'Ele_Order_Num': '9',
#                      'Ele_Company_Name_Selected': '1', 'Ele_Revenue_Selected': '1', 'Ele_Fapiao_No_Selected': '0',
#                      'Ele_Invoice_Name': 'CS + Company Name + Invoice No + Revenue',
#                      'Ele_Invoice_Files_Import_URL': 'N:\\Company Data\\FCO\\11.全电发票',
#                      'Ele_Invoice_Files_Export_URL': 'N:\\XM Softlines\\1. Project\\3. Finance\\02. WIP\\全电发票 2023\\10',
#                      '名称': '编号', 'Chen, Iris': '6375287',
#                      'Chen, Eunice': '6375162',
#                      'Ding, Daisy': '6160431', 'Du, Miley': '6375211', 'Guan, Elaine': '6375125',
#                      'Huang, Mary': '6375104', 'Jiao, Joyce': '6375079', 'Lai, Tailor': '6375014',
#                      'Lao, Keely': '6375134', 'Lin, Tina': '6375091', 'Lv, Rita': '6375135', 'Ma, Ella': '6160372',
#                      'Qiu, Dora': '6375241', 'Qiu, Emily': '6375235', 'Shen, Jewel': '6375124', 'Weng, Cora': '6375134',
#                      'Yang, Stacey': '6375142', 'Zhang, Judy': '6375176', 'Zhang, Wendy': '6375210',
#                      'Zhuo, Mia': '6375260', 'Huang, Holly': '6375162', 'Li, Cathy': '6375166', 'Yeh, Lynne': '6375134',
#                      'Zhang, Lyndon': '6375294', 'Wu, Jemma': '6375134', 'Luo, Luca': '6160275',
#                      'Ruan, Nicole': '6375183', 'Zhou, Judith': '6160350', 'Gan, Jasper': '6160244',
#                      'Ma, Ada': '6160185', 'You, Sofia': '6375105', 'Su, Layla': '6160385', 'Yang, Beauty': '6375308',
#                      'Huang, May': '6160385', 'Chen, Claudia': '6375162', 'Cai, Barry': '6375313',
#                      'Gong, Joy': '6375176', 'chen, sarah': '6375312', 'Chen, Raney': '6375162', 'Pan, Peki': '6375201',
#                      'Liu, Amber': '6375342', 'Chen, Kate': '6375337', 'Liu, Mia': '6375162', 'Liu, Morita': '6375336',
#                      'Peng, Penny': '6375351', 'Zhang, Alaia': '6375350', 'Huang, Even': '6375359',
#                      'Lin, Linda': '6375134', 'Lu, Joanna': '6375347', 'Wei, Wynne': '6375358',
#                      'Chen, Sarah': '6375312', 'Chen, Nemo': '6160291',
#                      'Xu, Jimmy': '6160343',
#                      'Su, Lucky': '6181557', 'Dai, Jocelyn': '6375017', 'Yang, Alisa': '6375038',
#                      'Zou, Rudi': '6375039', 'Wang, Carry': '6375064', 'Zhang, Lynn': '6375089', 'Wu, Alan': '6375092',
#                      'Li, Jesse': '6375093', 'Ou, Ida': '6375112', 'Miao, Molly': '6375158', 'Ye, Anne': '6375182',
#                      'Zeng, Cris': '6375184', 'Lin, Jenny': '6375252', 'Lin, Lucy': '6375253', 'Chen, Limi': '6375275',
#                      'Chen, Nikki': '6375277', 'Ye, Carter': '6375279', 'Wu, Mindy': '6375286', 'Han, Amy': '6375299',
#                      'Shen, Rocy': '6375302', 'Chen, Bella': '6375304', 'Ke, Coco': '6375314', 'Chen, Helen': '6375326',
#                      'Huang, Edwina': '6375330', 'Ma, Even': '6375331', 'Zhong, Teddy': '6375023',
#                      'Ou, Yedda': '6375024',
#                      'Zhang, Cathy': '6375043', 'Yang, Trison': '6375062', 'Huang, Moon': '6375084',
#                      'Qin, Bruce': '6375119', 'Zheng, Damon': '6375122', 'Ye, Valentine': '6375150',
#                      'Zhang, Dragon': '6375177', 'Zheng, Ariel': '6375196', 'Lu, Esther': '6375231',
#                      'Yang, Miya': '6375249', 'Zhan, Milla': '6375271', 'Lv, Linda': '6375273', 'Zeng, Tim': '6375280',
#                      'Xu, Simba': '6375282', 'Wang, Peter': '6375292', 'Zhou, Sean': '6375306',
#                      'Zeng, Winnie': '6375320', 'Chen, Echo': '6375321', 'Yu, Coley': '6375323',
#                      'Chen, Leah': '6375324', 'Ji, Sunny': '6375329', 'Li, Roy': '6375339', 'Liu, Josie': '6375341',
#                      'Zhang, Yvette': '6375349', 'Lin, Charlotte': '6375354', 'Pan, James': '6375355',
#                      'Yan, Alex': '6375356', 'Lin, Carl': '6375360', 'Xiao, Dennis': '6375362',
#                      'Cheng, Ethan': '6375369', 'Chen, Jacy': '6375372'}
#     staff_dict = {
#         'CHM': ['Chen, Nemo', 'Xu, Jimmy', 'Su, Lucky', 'Dai, Jocelyn', 'Yang, Alisa', 'Zou, Rudi',
#                 'Wang, Carry', 'Zhang, Lynn', 'Wu, Alan', 'Li, Jesse', 'Ou, Ida', 'Miao, Molly',
#                 'Ye, Anne', 'Zeng, Cris', 'Lin, Jenny', 'Lin, Lucy', 'Chen, Limi', 'Chen, Nikki',
#                 'Ye, Carter', 'Wu, Mindy', 'Han, Amy', 'Shen, Rocy', 'Chen, Bella', 'Ke, Coco',
#                 'Chen, Helen', 'Huang, Edwina', 'Ma, Even'],
#         'PHY': ['Zhong, Teddy', 'Ou, Yedda', 'Zhang, Cathy', 'Yang, Trison', 'Huang, Moon', 'Qin, Bruce',
#                 'Zheng, Damon', 'Ye, Valentine', 'Zhang, Dragon', 'Zheng, Ariel', 'Lu, Esther',
#                 'Yang, Miya', 'Zhan, Milla', 'Lv, Linda', 'Zeng, Tim', 'Xu, Simba', 'Wang, Peter',
#                 'Zhou, Sean', 'Zeng, Winnie', 'Chen, Echo', 'Yu, Coley', 'Chen, Leah', 'Ji, Sunny',
#                 'Li, Roy', 'Liu, Josie', 'Zhang, Yvette', 'Lin, Charlotte', 'Pan, James', 'Yan, Alex',
#                 'Lin, Carl', 'Xiao, Dennis', 'Cheng, Ethan', 'Chen, Jacy'],
#         'CS': ['Chen, Iris', 'Chen, Eunice', 'Ding, Daisy', 'Du, Miley', 'Guan, Elaine', 'Huang, Mary',
#                'Jiao, Joyce', 'Lai, Tailor', 'Lao, Keely', 'Lin, Tina', 'Lv, Rita', 'Ma, Ella',
#                'Qiu, Dora', 'Qiu, Emily', 'Shen, Jewel', 'Weng, Cora', 'Yang, Stacey', 'Zhang, Judy',
#                'Zhang, Wendy', 'Zhuo, Mia', 'Huang, Holly', 'Li, Cathy', 'Yeh, Lynne', 'Zhang, Lyndon',
#                'Wu, Jemma', 'Su, Layla', 'Yang, Beauty', 'Huang, May', 'Chen, Claudia', 'Cai, Barry',
#                'Gong, Joy', 'chen, sarah', 'Chen, Raney', 'Pan, Peki', 'Liu, Amber', 'Chen, Kate',
#                'Liu, Mia', 'Liu, Morita', 'Peng, Penny', 'Zhang, Alaia', 'Huang, Even', 'Lin, Linda',
#                'Lu, Joanna', 'Wei, Wynne', 'Chen, Sarah'],
#     }
#     allocator = RevenueAllocator()
#
#     # 定义CSV文件的表头
#     res_headers = ['order_no', 'material_code', 'item', 'dept', 'dept_revenue', 'dept_hours']
#     res2_headers = ['order_no', 'material_code', 'item', 'dept', 'dept_revenue', 'dept_hours',
#                    'allocated_date', 'allocated_hours', 'staff_name']
#
#     # 第一次写入时包含表头
#     first_write = True
#
#     for revenueData in revenueDatas:
#         res = allocator.calculate_revenue_allocation(revenueData, configContent)
#         res2 = allocator.allocate_working_hours(res, 8, datetime.date(2025, 4, 1), datetime.date(2025, 4, 30),
#                                      staff_dict)
#         res_df = pd.DataFrame(res)
#         res_df2 = pd.DataFrame(res2)
#
#         # 写入CSV文件，第一次写入时包含表头
#         res_df.to_csv('res.csv', index=False, mode='a', header=first_write)
#         res_df2.to_csv('res2.csv', index=False, mode='a', header=first_write)
#
#         # 第一次写入后设置为False
#         first_write = False
#
#     print(res)
