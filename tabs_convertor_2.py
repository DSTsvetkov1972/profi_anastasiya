import package.init_project
import os, sys
sys.path.append(os.getcwd())

from package.fns import get_source, get_result, format_res_file
from colorama import Fore
from datetime import datetime


while True:

    try:
        os.startfile('исходник.xlsx')
        print(Fore.YELLOW + 'Скопируйте данные в файл "исходник.xlsx"\nзакройте его и нажмите ввод' + Fore.RESET)
        input()

        res_file = f'результат_{str(datetime.now())[:19].replace(':', '-')}.xlsx'
        source_df = get_source()
        result_df = get_result(source_df) 
        result_df.to_excel(res_file, index = None, header=None)
        format_res_file(res_file)
        
        print(Fore.GREEN + f"Рассчёты выполнены успешно и сохранены в файл:\n{res_file}" + Fore.RESET)
        
        os.startfile(res_file)
    except Exception as e:
        print(Fore.RED + repr(e) + Fore.RESET)

    continue