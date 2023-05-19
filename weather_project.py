from bs4 import BeautifulSoup
import requests
import pandas as pd

def weathercom():

    html_text = requests.get('https://weather.com/weather/today/l/cd5287a9b7c6082e70de21891369ada61322a7aa1fc24cca26803dcbd75c78d9').text
    soup = BeautifulSoup(html_text, 'lxml')
    main = soup.find('main', class_ = 'region-main regionMain DaybreakLargeScreen--regionMain--1FzNI')

    def currenttemp():


        location = main.find('h1', class_ = 'CurrentConditions--location--1YWj_').text
        temperature = main.find('div', class_ = 'CurrentConditions--primary--2DOqs').span.text

        print(f'''
        {location}
        {temperature}
        ''')


    def forecast():
        forecast = main.find('div', class_ = 'DailyWeatherCard--TableWrapper--2bB37')
        daily_forecast = forecast.find_all('li', class_ = 'Column--column--3tAuz')
        for day in daily_forecast:
            title = day.find('h3', class_ = 'Column--label--2s30x Column--default--2-Kpw').span.text
            temp = day.find('div', class_ = 'Column--temp--1sO_J').span.text
            print(f'''
            {title}
            {temp}
            ''')

    def details():
        today_details = main.find('div', 'TodayDetailsCard--detailsContainer--2yLtL')
        details = today_details.find_all('div', class_ = 'ListItem--listItem--25ojW WeatherDetailsListItem--WeatherDetailsListItem--1CnRC')
        for detail in details:

            title = detail.find('div', class_ = 'WeatherDetailsListItem--label--2ZacS').text
            value = detail.find('div', class_ = 'WeatherDetailsListItem--wxData--kK35q').text


            print(f'''
            {title}
            {value}
            ''')

    currenttemp()
    details()
    forecast()

def export(data):
    df = pd.DataFrame(data)
    df.to_excel('weather.xlsx')

if __name__ == '__main__':
    data = weathercom()
    export(data)
    print('Done')












    






