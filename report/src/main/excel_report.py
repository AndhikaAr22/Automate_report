from report.src.operators.xlsx_report_plugin import ExcelReportPlugin
from report.src.utils_discord.discord_webhook import send_to_discord
import os
import json

base_path = os.sep.join(os.getcwd().split(os.sep)[:-3])
print(f'base path = {base_path}')

input_file = base_path + '/input_data/supermarket_sales.xlsx'
output_file = base_path + '/output_data/new_report_penjualan_supermarket1.xlsx'

# json file
configs = open(base_path + '/configs/webhook.json')
webhook_url = json.load(configs)['webhook_url']




automate = ExcelReportPlugin(
    input_file=input_file,
    output_file=output_file

)

if __name__ == "__main__":
        automate.main()
        send_to_discord(webhook_url,output_file)

