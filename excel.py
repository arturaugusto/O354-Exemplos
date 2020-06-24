from O365 import Account

credentials = ('XXXXXXXXXXX', 'XXXXXXX')

account = Account(credentials, tenant_id='XXXXXXXX')
if account.authenticate(scopes=['basic', 'message_all']):
    print('Authenticated!')

storage = account.storage()

drives = storage.get_drives()

my_drive = drives[0]

root_folder = my_drive.get_root_folder()

for x in root_folder.get_items(limit=25):
    print(x)

from O365.excel import WorkBook

files = my_drive.search('Pasta de trabalho 1', limit=1)

f = next(files)

excel_file = WorkBook(f)

ws = excel_file.get_worksheet('Planilha1')
cella1 = ws.get_range('A1')
cella1.values = 'Olá IPT'
cella1.update()
