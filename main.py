from os import listdir
from os.path import isfile, join
from robot import Robot
from data_structures import Data


def main() -> None:
    excel_path = r'\\dbu00234\c$\Temp\Сверка'
    excel_name = [f for f in listdir(excel_path) if isfile(join(excel_path, f))][0]

    data = Data(
        usr='robot',
        psw='Asd_24-08-2022',
        process_name='COLVIR',
        process_path=r'C:\CBS_R_NEW\CBS_R\COLVIR.EXE',
        excel_path=excel_path,
        excel_name=excel_name,
        zip_dir=r'C:\Reports',
        zip_file=r'C:\Users\robot.ad\Desktop\reports.zip',
        email_list=['zhekenova.a@otbasybank.kz', 'abdullayeva.b@otbasybank.kz', 'baktibay.d@otbasybank.kz', 'abdieva.g@otbasybank.kz', 'robot.ad@hcsbkkz.loc'],
    )
    robot = Robot(**data.data)
    robot.run()


if __name__ == '__main__':
    main()

# try:
#     data = get_input_data().data
#     robot = Robot(**data)
#     robot.run()
# except IndexError:
#     today = dt.now().strftime('%d.%m.%Y')
#     email = Email(
#         email_info=EmailInfo(email_list=ast.literal_eval({{email_list}})),
#         subject=f'Отсутствует файл для проверки {today}',
#         body=f'Отсутствует файл для проверки в \\\\dbu00234\\c$\\Temp\\Сверка\n\n'
#                 f'Пожалуйста добавьте его перед запуском робота.\n\nСообщение сгенерировано автоматически. '
#                 f'Просьба не отвечать.',
#     )
#     email.run()
