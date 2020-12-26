import glob
import re
import yaml
import openpyxl
import configparser

if __name__ == '__main__':
    config_ini = configparser.ConfigParser()
    config_ini.read('config.ini', encoding='utf-8')

    path = config_ini['DEFAULT']['Path']
    wb = openpyxl.load_workbook('weapons.xlsx')
    sheet = wb.active
    row_num = 3  # 先頭2行はヘッダー

    file_list = glob.glob(path, recursive=True)
    for r in file_list:
        with open(r, 'r', encoding='utf_8') as file:
            y = yaml.safe_load(file)
            name = next(iter(y))  # AK-47とか根を取得する

            sheet['A' + str(row_num)].value = row_num - 2

            root = y[name]
            sheet['B' + str(row_num)].value = name

            item_information = root['Item_Information']
            if item_information:
                sheet['C' + str(row_num)].value = re.sub('&.', '', item_information.get('Item_Name', '-'))
                sheet['D' + str(row_num)].value = item_information.get('Inventory_Control', '-')

            shooting = root.get('Shooting')
            if shooting:
                sheet['E' + str(row_num)].value = shooting.get('Projectile_Amount', 0)
                sheet['F' + str(row_num)].value = shooting.get('Projectile_Type', '-')
                sheet['G' + str(row_num)].value = shooting.get('Projectile_Speed', 0)
                sheet['H' + str(row_num)].value = shooting.get('Projectile_Damage', 0)
                sheet['I' + str(row_num)].value = shooting.get('Bullet_Spread', 0.0)

            fully_automatic = root.get('Fully_Automatic')
            if fully_automatic:
                sheet['J' + str(row_num)].value = fully_automatic.get('Fire_Rate', 0)

            reload = root.get('Reload')
            if reload:
                sheet['K' + str(row_num)].value = reload.get('Reload_Amount', 0)
                sheet['L' + str(row_num)].value = reload.get('Reload_Duration', 0)

            firearm_action = root.get('Firearm_Action')
            if firearm_action:
                sheet['M' + str(row_num)].value = firearm_action.get('Type', '-')
                sheet['N' + str(row_num)].value = firearm_action.get('Close_Duration', 0)
                sheet['O' + str(row_num)].value = firearm_action.get('Close_Shoot_Delay', 0)

            scope = root.get('Scope')
            if scope:
                sheet['P' + str(row_num)].value = scope.get('Zoom_Amount', 0)
                sheet['Q' + str(row_num)].value = scope.get('Zoom_Bullet_Spread', 0.0)

            headshot = root.get('Headshot')
            if headshot:
                sheet['R' + str(row_num)].value = headshot.get('Bonus_Damage', 0.0)
        row_num += 1
    wb.save('weapons_result.xlsx')
