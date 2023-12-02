from xlsxwriter import Workbook
from requests import get
from fake_useragent import UserAgent


def get_tx_count(address: str) -> int:

    ua = UserAgent(os=["windows"], browsers=['chrome'])

    headers = {
        'authority': 'layerzeroscan.com',
        'accept': '*/*',
        'accept-language': 'ru,en;q=0.9',
        'content-type': 'application/json',
        'referer': f'https://layerzeroscan.com/address/{address}',
        'sec-ch-ua': '"Chromium";v="118", "YaBrowser";v="23", "Not=A?Brand";v="99"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-origin',
        'user-agent': ua.random,
        'x-kl-kis-ajax-request': 'Ajax_Request',
    }

    params = {
        'input': '{"stage":"mainnet","address":"' + address + '"}',
    }

    response = get('https://layerzeroscan.com/api/trpc/metrics.volume', params=params, headers=headers)

    data =  response.json()

    return data['result']['data']['volumeTotal']


def xls_constructor(columns: list[int, str, int]) -> None:
    pretty_columns = [
        "Id",
        "Address",
        "Tx Count"
    ]
    
    workbook = Workbook('LayerZero Stats.xlsx')
    worksheet = workbook.add_worksheet("Stats")

    header_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'border': 1
    })

    for col_num, column in enumerate(pretty_columns):
        worksheet.write(0, col_num, column, header_format)

    for row_num, data in enumerate(columns, start=1):
        for col_num, info in enumerate(data):
            worksheet.write(row_num, col_num, info)

    row_format = workbook.add_format({'align': 'center'})

    for id, size in enumerate([5, 45, 9]):
        worksheet.set_column(id, id, size, row_format)

    workbook.close()


def main(addresses: list[str, str]) -> None:
    data = []

    for id, address in enumerate(addresses, start=1):
        tx_count = get_tx_count(address)

        data.append([id, address, tx_count])

        print(f'{id} | Address: {address} | Number of transactions: {tx_count}')

    xls_constructor(data)


if __name__ == '__main__':

    with open('addresses.txt') as file:
        addresses = [i.strip() for i in file]

    main(addresses)
