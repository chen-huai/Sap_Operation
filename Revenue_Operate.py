import datetime
from chinese_calendar import is_holiday
import pandas as pd
import os

# 新增计算方法
class RevenueAllocator:
    def __init__(self):
        self.hours_data = pd.DataFrame(columns=['date', 'staff_name', 'hours', 'order_no', 'dept', 'week'])
        self.hours_file = None  # 将在使用时根据日期确定文件名
        self.unallocated_hours_file = None  # 未分配工时记录文件

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
                # 确保week列存在
                if 'week' not in self.hours_data.columns:
                    self.hours_data['week'] = self.hours_data['date'].apply(self._get_week_number)
            except Exception as e:
                print(f"Error loading hours data: {e}")
                self.hours_data = pd.DataFrame(columns=['date', 'staff_name', 'hours', 'order_no', 'dept', 'week'])

    def _save_hours_data(self):
        """
        保存工时数据到对应月份的文件
        """
        if self.hours_file is None:
            return
            
        try:
            # 确保week列存在
            if 'week' not in self.hours_data.columns:
                self.hours_data['week'] = self.hours_data['date'].apply(self._get_week_number)
            
            self.hours_data.to_csv(self.hours_file, index=False)
        except Exception as e:
            print(f"Error saving hours data: {e}")

    def _get_staff_daily_hours(self, date, staff_name=None):
        """
        获取指定日期的工作时长
        :param date: 日期
        :param staff_name: 员工姓名，如果为None则返回所有员工的工作时长
        :return: 如果指定员工，返回该员工的工作时长；如果未指定员工，返回所有员工的工作时长字典
        """
        if self.hours_data.empty:
            return {} if staff_name is None else 0
        
        if staff_name is not None:
            mask = (self.hours_data['date'] == date) & (self.hours_data['staff_name'] == staff_name)
            return self.hours_data.loc[mask, 'hours'].sum()
        else:
            mask = self.hours_data['date'] == date
            return self.hours_data.loc[mask].groupby('staff_name')['hours'].sum().to_dict()

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
            'dept': [dept],
            'week': [self._get_week_number(date)]
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

    def allocate_department_hours(self, revenueData, configContent):
        """动态配置的收入分配计算方法"""
        material_code = revenueData.get('Material Code', '')
        base = (float(revenueData['Revenue']) - float(revenueData['Total Subcon Cost']) / 1.06) * float(
            configContent.get('Plan_Cost_Parameter'))
        primary_cs = revenueData.get('Primary CS', '')  # 获取Primary CS字段

        # 获取有效位数配置
        significant_digits = int(configContent.get('Significant_Digits', 0))

        result = {
            'business_dept_1000_revenue': 0, 'lab_1000_revenue': 0, 'business_dept_1000_hours': 0, 'lab_1000_hours': 0,
            'business_dept_2000_revenue': 0, 'lab_2000_revenue': 0, 'business_dept_2000_hours': 0, 'lab_2000_hours': 0,
            'lab_1000': '', 'lab_2000': '', 'order_no': revenueData.get('Order Number', ''),
            'material_code_1000': '',
            'material_code_2000': '',
            'primary_cs': primary_cs  # 添加Primary CS字段
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

            # 计算并格式化金额和工时
            business_dept_1000_revenue = round(base * (1 - lab_cost), 2)
            lab_1000_revenue = round(base * lab_cost, 2)
            business_dept_1000_hours = round((base * (1 - lab_cost)) / business_dept_rate, significant_digits)
            lab_1000_hours = round((base * lab_cost) / lab_rate, significant_digits)

            result.update({
                'business_dept_1000_revenue': business_dept_1000_revenue,
                'lab_1000_revenue': lab_1000_revenue,
                'business_dept_1000_hours': business_dept_1000_hours,
                'lab_1000_hours': lab_1000_hours,
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

            # 计算并格式化金额和工时
            business_dept_1000_revenue = round(base * proportion_1000 * (1 - lab_1000_cost), 2)
            lab_1000_revenue = round(base * proportion_1000 * lab_1000_cost, 2)
            business_dept_2000_revenue = round(base * proportion_2000 * (1 - lab_2000_cost), 2)
            lab_2000_revenue = round(base * proportion_2000 * lab_2000_cost, 2)

            business_dept_1000_hours = round((base * proportion_1000 * (1 - lab_1000_cost)) / business_dept_rate,
                                             significant_digits)
            lab_1000_hours = round((base * proportion_1000 * lab_1000_cost) / lab_1000_rate, significant_digits)
            business_dept_2000_hours = round((base * proportion_2000 * (1 - lab_2000_cost)) / business_dept_rate,
                                             significant_digits)
            lab_2000_hours = round((base * proportion_2000 * lab_2000_cost) / lab_2000_rate, significant_digits)

            result.update({
                'business_dept_1000_revenue': business_dept_1000_revenue,
                'lab_1000_revenue': lab_1000_revenue,
                'business_dept_2000_revenue': business_dept_2000_revenue,
                'lab_2000_revenue': lab_2000_revenue,
                'business_dept_1000_hours': business_dept_1000_hours,
                'lab_1000_hours': lab_1000_hours,
                'business_dept_2000_hours': business_dept_2000_hours,
                'lab_2000_hours': lab_2000_hours,
                'lab_1000': lab_1000,
                'lab_2000': lab_2000,
                'material_code_1000': configContent.get(f"{material_code}_mc").split('/')[0],
                'material_code_2000': configContent.get(f"{material_code}_mc").split('/')[1]
            })

        # 在返回结果前增加数据结构整理
        return [
            {  # 1000业务部门
                'order_no': revenueData['Order Number'],
                'material_code': result.get("material_code_1000"),
                'item': '1000',
                'dept': business_dept,
                'dept_revenue': round(result['business_dept_1000_revenue'], 2),
                'dept_hours': round(result['business_dept_1000_hours'], significant_digits),
                'original_hours': round(result['business_dept_1000_hours'], significant_digits),
                'primary_cs': result['primary_cs']  # 添加Primary CS字段
            },
            {  # 1000实验室
                'order_no': revenueData['Order Number'],
                'material_code': result.get("material_code_1000"),
                'item': '1000',
                'dept': result['lab_1000'],
                'dept_revenue': round(result['lab_1000_revenue'], 2),
                'dept_hours': round(result['lab_1000_hours'], significant_digits),
                'original_hours': round(result['lab_1000_hours'], significant_digits),
                'primary_cs': result['primary_cs']  # 添加Primary CS字段
            },
            {  # 2000业务部门
                'order_no': revenueData['Order Number'],
                'material_code': result.get("material_code_2000"),
                'item': '2000',
                'dept': business_dept,
                'dept_revenue': round(result['business_dept_2000_revenue'], 2),
                'dept_hours': round(result['business_dept_2000_hours'], significant_digits),
                'original_hours': round(result['business_dept_2000_hours'], significant_digits),
                'primary_cs': result['primary_cs']  # 添加Primary CS字段
            },
            {  # 2000实验室
                'order_no': revenueData['Order Number'],
                'material_code': result.get("material_code_2000"),
                'item': '2000',
                'dept': result['lab_2000'],
                'dept_revenue': round(result['lab_2000_revenue'], 2),
                'dept_hours': round(result['lab_2000_hours'], significant_digits),
                'original_hours': round(result['lab_2000_hours'], significant_digits),
                'primary_cs': result['primary_cs']  # 添加Primary CS字段
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

    def _save_unallocated_hours(self, unallocated_data, export_path):
        """
        保存未分配工时信息到Excel文件
        :param unallocated_data: 未分配工时数据列表
        :param export_path: Excel文件保存路径
        """
        if not unallocated_data:
            return

        # 创建DataFrame
        df = pd.DataFrame(unallocated_data)
        
        # 设置文件名和路径
        current_date = datetime.datetime.now().strftime("%Y%m%d")
        file_name = f"unallocated_hours_{current_date}.xlsx"
        self.unallocated_hours_file = os.path.join(export_path, file_name)
        
        try:
            # 如果文件已存在，追加数据
            if os.path.exists(self.unallocated_hours_file):
                existing_df = pd.read_excel(self.unallocated_hours_file)
                df = pd.concat([existing_df, df], ignore_index=True)
            
            # 保存到Excel
            df.to_excel(self.unallocated_hours_file, index=False)
            print(f"Unallocated hours information saved to {self.unallocated_hours_file}")
        except Exception as e:
            print(f"Error saving unallocated hours: {e}")

    def _get_weekly_records_count(self, staff_name, week_number):
        """
        获取指定员工在指定周数的记录数量
        :param staff_name: 员工姓名
        :param week_number: 周数
        :return: 记录数量
        """
        if self.hours_data.empty:
            return 0
        
        # 确保week列存在
        if 'week' not in self.hours_data.columns:
            self.hours_data['week'] = self.hours_data['date'].apply(self._get_week_number)
        
        mask = (self.hours_data['staff_name'] == staff_name) & (self.hours_data['week'] == week_number)
        return len(self.hours_data[mask])

    def allocate_person_hours(self, results, max_hours_per_day, start_date, end_date, staff_dict, configContent):
        """
        工时分配方法
        :param results: 需要分配工时的记录列表
        :param max_hours_per_day: 每人每天最大工作时长
        :param start_date: 开始日期
        :param end_date: 结束日期
        :param staff_dict: 部门人员字典 {部门: [员工编号1, 员工编号2...]}
        :return: 分配后的工时记录列表
        """
        # 获取有效位数配置
        significant_digits = int(configContent.get('Significant_Digits', 0))

        # 加载现有工时数据
        self._load_hours_data(start_date, configContent.get('Hour_Files_Export_URL'))

        # 过滤零工时数据
        filtered = [r for r in results if r['dept_hours'] > 0]

        # 生成工作日历
        work_days = self.generate_work_days(start_date, end_date)
        if not work_days:
            return []

        final_results = []
        unallocated_data = []  # 存储未分配工时信息

        # 按订单号和部门分组处理工时
        order_dept_groups = {}
        for record in filtered:
            key = (record['order_no'], record['dept'])
            if key not in order_dept_groups:
                order_dept_groups[key] = []
            order_dept_groups[key].append(record)

        # 创建一个全局的工作日分配计数器
        global_workday_counter = {day: 0 for day in work_days}

        # 处理每个订单组
        for (order_no, dept), order_records in order_dept_groups.items():
            # 获取对应部门的员工列表
            staff_list = staff_dict.get(dept, [])
            if not staff_list:
                print(f"Warning: No staff found for department {dept}")
                continue

            # 获取Primary CS（如果存在）
            primary_cs = order_records[0].get('primary_cs', '')
            primary_cs_allocated = False  # 标记Primary CS是否已分配

            # 计算总工时
            total_hours = sum(record['dept_hours'] for record in order_records)
            print(f"Processing order {order_no} in department {dept} with total hours: {total_hours}")

            # 计算每天需要分配的人数（向上取整）
            staff_per_day = max(1, int(total_hours / (max_hours_per_day * len(work_days))) + 1)
            staff_per_day = min(staff_per_day, len(staff_list))

            # 计算每个工作日应该分配的平均记录数
            total_records = len(order_records)
            avg_records_per_day = total_records / len(work_days)
            
            # 将工作日按分配次数排序，优先选择分配次数较少的工作日
            sorted_work_days = sorted(work_days, key=lambda x: global_workday_counter[x])
            
            # 如果是CS部门且有Primary CS，优先分配一次给Primary CS
            if dept == 'CS' and primary_cs and primary_cs in staff_list:
                for work_day in sorted_work_days:
                    # 获取当天已分配工时
                    daily_allocations = self._get_staff_daily_hours(work_day)
                    
                    # 检查Primary CS当天是否还有可用工时
                    primary_cs_hours = daily_allocations.get(primary_cs, 0)
                    if primary_cs_hours < max_hours_per_day:
                        # 计算可分配的工时
                        available_hours = max_hours_per_day - primary_cs_hours
                        allocated_hours = min(available_hours, total_hours)
                        
                        # 更新工时记录
                        self._update_staff_daily_hours(work_day, primary_cs, allocated_hours, order_no, dept)
                        
                        # 增加工作日的分配计数
                        global_workday_counter[work_day] += 1
                        
                        # 标记Primary CS已分配
                        primary_cs_allocated = True
                        
                        # 添加分配记录
                        for record in order_records:
                            if record['dept_hours'] <= 0:
                                continue

                            # 计算该记录可分配的工时
                            record_hours = min(allocated_hours, record['dept_hours'])
                            if record_hours <= 0:
                                continue

                            new_record = record.copy()
                            new_record.update({
                                'allocated_date': work_day,
                                'allocated_day': work_day.day,
                                'allocated_hours': round(record_hours, significant_digits),
                                'staff_name': primary_cs,
                                'staff_id': configContent.get(primary_cs),
                                'week': self._get_week_number(work_day)
                            })
                            final_results.append(new_record)

                            record['dept_hours'] -= record_hours
                            allocated_hours -= record_hours
                            total_hours -= record_hours

                            if allocated_hours <= 0:
                                break
                        
                        # 如果Primary CS分配完成，跳出循环
                        break

            # 遍历每个工作日进行分配
            for work_day in sorted_work_days:
                # 如果当前工作日的分配次数已经超过平均值，跳过
                if global_workday_counter[work_day] > avg_records_per_day * 1.1:  # 允许10%的浮动
                    continue

                # 获取当天已分配工时
                daily_allocations = self._get_staff_daily_hours(work_day)
                
                # 过滤出该部门的已分配人员
                dept_allocations = {name: hours for name, hours in daily_allocations.items() 
                                  if name in staff_list}
                
                # 获取可用的员工（按工时从少到多排序）
                available_staff = sorted(
                    [staff for staff in staff_list 
                     if staff not in dept_allocations or dept_allocations[staff] < max_hours_per_day],
                    key=lambda x: dept_allocations.get(x, 0)
                )
                
                if not available_staff:
                    continue

                # 计算当天需要分配的总工时
                remaining_hours = sum(record['dept_hours'] for record in order_records)
                if remaining_hours <= 0:
                    break

                # 计算每个员工应分配的工时
                hours_per_staff = min(
                    max_hours_per_day,
                    remaining_hours / min(len(available_staff), staff_per_day)
                )

                # 获取当前周数
                current_week = self._get_week_number(work_day)

                # 分配工时给员工
                for staff_name in available_staff[:staff_per_day]:
                    if remaining_hours <= 0:
                        break

                    # 检查该员工本周记录数是否已达到上限
                    weekly_records = self._get_weekly_records_count(staff_name, current_week)
                    if weekly_records >= 14:
                        continue

                    # 计算该员工可分配的最大工时
                    available_hours = max_hours_per_day - dept_allocations.get(staff_name, 0)
                    if available_hours <= 0:
                        continue

                    # 计算实际分配的工时
                    allocated_hours = min(hours_per_staff, available_hours, remaining_hours)
                    if allocated_hours <= 0:
                        continue

                    # 更新工时记录
                    self._update_staff_daily_hours(work_day, staff_name, allocated_hours, order_no, dept)
                    
                    # 增加工作日的分配计数
                    global_workday_counter[work_day] += 1
                    
                    # 添加分配记录
                    for record in order_records:
                        if record['dept_hours'] <= 0:
                            continue

                        # 计算该记录可分配的工时
                        record_hours = min(allocated_hours, record['dept_hours'])
                        if record_hours <= 0:
                            continue

                        new_record = record.copy()
                        new_record.update({
                            'allocated_date': work_day,
                            'allocated_day': work_day.day,
                            'allocated_hours': round(record_hours, significant_digits),
                            'staff_name': staff_name,
                            'staff_id': configContent.get(staff_name),
                            'week': current_week
                        })
                        final_results.append(new_record)

                        record['dept_hours'] -= record_hours
                        allocated_hours -= record_hours
                        remaining_hours -= record_hours

                        if allocated_hours <= 0:
                            break

                    if remaining_hours <= 0:
                        break

            # 检查是否所有工时都已分配
            remaining_total = sum(record['dept_hours'] for record in order_records)
            if remaining_total > 0:
                # 检查所有工作日是否都已分配满8小时
                all_days_full = True
                for work_day in work_days:
                    daily_allocations = self._get_staff_daily_hours(work_day)
                    dept_allocations = {name: hours for name, hours in daily_allocations.items() 
                                      if name in staff_list}
                    if any(hours < max_hours_per_day for hours in dept_allocations.values()):
                        all_days_full = False
                        break

                # 只有当所有工作日都分配满8小时时，才记录未分配工时
                if all_days_full:
                    for record in order_records:
                        if record['dept_hours'] > 0:
                            unallocated_data.append({
                                'order_no': order_no,
                                'dept': dept,
                                'material_code': record['material_code'],
                                'item': record['item'],
                                'remaining_hours': round(record['dept_hours'], significant_digits),
                                'original_hours': record['original_hours'],
                                'check_date': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            })

        # 保存未分配工时信息到指定路径
        self._save_unallocated_hours(unallocated_data, configContent.get('Hour_Files_Export_URL'))

        return final_results


# if __name__ == "__main__":
#     # 示例数据
#
#     revenueDatas = [
#         {'Order Number': 'ORD123',
#          'Material Code': 'T20-430-A2',
#          'Revenue': 20000,
#          'Total Subcon Cost': 2000},
#         {'Order Number': 'ORD124',
#          'Material Code': 'T75-405-A2',
#          'Revenue': 10000,
#          'Total Subcon Cost': 2000},
#         {'Order Number': 'ORD125',
#          'Material Code': 'T75-441-A2',
#          'Revenue': 30000,
#          'Total Subcon Cost': 2000},
#         {'Order Number': 'ORD126',
#          'Material Code': 'T20-441-00',
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
#                      'Cheng, Ethan': '6375369', 'Chen, Jacy': '6375372', 'Hour_Files_Export_URL': "N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\2.财务\\2.SAP\\1.ODM Data - XM\\3.Hours"}
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
#         res = allocator.allocate_department_hours(revenueData, configContent)
#         res2 = allocator.allocate_person_hours(res, 8, datetime.date(2025, 4, 1), datetime.date(2025, 4, 30),
#                                      staff_dict,configContent)
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
