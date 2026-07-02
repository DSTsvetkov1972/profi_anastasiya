import package.init_project
import os, sys
sys.path.append(os.getcwd())

from package.fns import get_source, get_result, format_res_file
from colorama import Fore
from datetime import datetime

os.startfile('исходник.xlsx')

while True:

    try:

        print(Fore.YELLOW + '\nСкопируйте данные в файл "исходник.xlsx" и нажмите ввод' + Fore.RESET)
        input()

        res_file = f'результат_{str(datetime.now())[:19].replace(':', '-')}.xlsx'
        source_df = get_source()
        result_df = get_result(source_df)
        # print(result_df)
        result_df.to_excel(res_file, index = None, header=None)
        format_res_file(res_file)
        
        print(Fore.GREEN + "Рассчёты выполнены успешно и сохранены в файл:" + Fore.RESET)
        print(Fore.BLACK+ res_file + Fore.RESET)
        print(Fore.CYAN + "(открыаем на рабочем столе)" + Fore.RESET)

        os.startfile(res_file)
    except Exception as e:
        print(Fore.RED + repr(e) + Fore.RESET)

    continue