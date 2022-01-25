import asyncio
from pyppeteer import launch
import openpyxl


async def main():
    browser = await launch({'headless': True, 'args': ['--no-sandbox']})
    page = await browser.newPage()

    await page.goto(
        'https://news.qq.com/zt2020/page/feiyan.htm?ADTAG=antipcell&isnm=1&qimei=8f312bab440eb3e9&devid=7bc8f7d110e0a931&shareto=wx#/?pool=jiangsu&ADTAG=antipcell')
    await asyncio.sleep(1)

    await page.click('.btn_fxdq')

    data_list = []
    # 获取高风险区数据
    # await asyncio.sleep(1)
    items_high = await page.xpath('//div[@class="content"]/div[1]/p')
    # print("高风险区数据：")
    for items_for_high in items_high:
        text_for_high = await (await items_for_high.getProperty('innerHTML')).jsonValue()
        data_list.append(text_for_high)
    # 获取中风险区数据
    # items_middle = await page.xpath('//div[@class="content"]/div[2]/p[position()>1]')
    items_middle = await page.xpath('//div[@class="content"]/div[2]/p')
    # print("中风险区数据：")
    for items_for_middle in items_middle:
        text_for_middle = await (await items_for_middle.getProperty('innerHTML')).jsonValue()
        data_list.append(text_for_middle)
    # 保存为xls格式
    path = r"C:\Users\16397\Desktop\中高风险地区疫情数据.xls"
    sheetStr = 'sheet1'
    write_to_excel(path, sheetStr, data_list)

    await browser.close()


async def wait_fornavigation(page, events):  # 等到某动作完成
    await asyncio.wait([
        events,
        page.waitForNavigation({'timeout': 50000}),
    ])


def write_to_excel(path, sheetStr, data):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = sheetStr

    count = 0

    for i in range(len(data)):
        sheet.cell(row=i + 1, column=1, value=data[i])
        if (data[i].find("郑州") != -1):
            count += 1
            print(data[i])
    print("中高风险区总计：{}个".format(count))

    workbook.save(path)


if __name__ == '__main__':
    asyncio.get_event_loop().run_until_complete(main())
