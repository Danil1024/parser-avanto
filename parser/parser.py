import requests
from bs4 import BeautifulSoup
import urllib3
import xlsxwriter

class Parser():
    def __init__(self) -> None:
        urllib3.disable_warnings()
        self.session = requests.Session()
        self.MAIN_URL = 'https://avantaauto.ru'
        self.HEADERS = {
                        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36'
                        }

    def get_all_url_brand(self) -> list:
        brand_list = list()
        response = self.session.get(url=self.MAIN_URL, headers=self.HEADERS,  verify=False)
        html_main_page = BeautifulSoup(response.text, 'lxml')
        all_brand_html = html_main_page.find('ul', class_='search-brand').find_all('a')
        for brang_html in all_brand_html:
            brand_url = self.MAIN_URL + brang_html.get('href')
            brand_list.append(brand_url)
        return brand_list
    
    def pars_brand_list(self, brand_list) -> list:
        cars_url_list = list()
        for brand_url in brand_list:
            cars = self.pars_brand_page(brand_url)
            cars_url_list += cars
        return cars_url_list
    
    def pars_brand_page(self, brand_url) -> list:
        cars = list()
        response = self.session.get(url=brand_url, headers=self.HEADERS, verify=False)
        html_brand_page = BeautifulSoup(response.text, 'lxml')
        cars_html = html_brand_page.find_all('div', class_='hit-card__title')
        for car_html in cars_html:
            car_url = self.MAIN_URL + car_html.find('a').get('href')
            cars.append(car_url)
        return cars
    
    def pars_cars_list(self, cars_url_list) -> list:
        cars_info_list = list()
        for car_url in cars_url_list:
            car_info = self.pars_car_page(car_url)
            if car_info == 'снята с продажи':
                continue
            cars_info_list.append(car_info)
        return cars_info_list
    

    def pars_car_page(self, car_url) -> list:
        print(car_url)
        response = self.session.get(url=car_url, headers=self.HEADERS, verify=False)
        html_car_page = BeautifulSoup(response.text, 'lxml')
        if html_car_page.find('span', class_='offer-main__price-actual text-center') is not None:
            print('машина снята с продажи')
            return 'снята с продажи'
        brand = html_car_page.find('div', id='bx_breadcrumb_2').find('span').text
        model = html_car_page.find('div', id='bx_breadcrumb_3').find('span').text
        full_name = html_car_page.find('h1', itemprop='name').text
        main_photo = self.MAIN_URL + html_car_page.find('img', class_='offer-main__car').get('src')
        car_old_price = html_car_page.find('span', class_='offer-main__price-crossed').find('span').text.replace(' ', '').replace('до', '')
        car_actual_price = html_car_page.find('span', class_='offer-main__price-actual').find('b').text.replace(' ', '')
        car_power = html_car_page.find('ul', class_='offer-specs').find_all('li', class_='offer-specs__item')[1]\
                                    .find('p', class_='offer-specs__item-accent').text.replace('\n', '').replace(' ', '')
        car_capacity = html_car_page.find('ul', class_='offer-specs').find_all('li', class_='offer-specs__item')[-1]\
                                    .find('p', class_='offer-specs__item-accent').text.replace('\n', '').replace(' ', '')

        equipments_2_options_string = str()
        equipments_2_html = html_car_page.find('div', class_='row slider_eq_and_price').find_all('ul', class_='diffs-card__data')
        for equipment_2_html in equipments_2_html:
            equipment_2_details_html = equipment_2_html.find_all('li')
            for equipment_2_detail_html in equipment_2_details_html:
                equipment_2_detail = equipment_2_detail_html.find('span').text
                equipments_2_options_string += equipment_2_detail + '|'
            equipments_2_options_string += '|'
        equipments_2_options_string = equipments_2_options_string[0:-2]

        equipments_name_string, equipments_power_string, equipments_kpp_string, equipments_fuel_string, equipment_old_price_string,\
        equipments_payment_per_month_string, equipments_benefit_string, equipments_special_credit_price_string, equipments_sum_price_string,\
        equipment_option_comfort_string, equipment_option_salon_string, equipment_option_visibility_string, equipment_option_safety_string,\
        equipment_option_theft_protection_string, equipment_option_multimedia_string, equipment_option_exterior_elements_string,\
        equipment_option_package_string, equipment_option_other_string = self.get_equipmets_info(html_car_page)
        engines_name_string, engines_kpp_string, engines_avto_class_string, engines_power_string, engines_capacity_string, engines_drive_unit_string,\
        engines_expenditure_string, engines_country_string, engines_engine_type_string, engines_fuel_string,\
        engines_amount_kpp_string = self.get_engines_info(html_car_page)
        color_name_string, color_code_string, color_url_string = self.get_color_info(html_car_page)
        return [brand, model, full_name, main_photo, car_actual_price, car_old_price, car_power, car_capacity, color_name_string, color_code_string, color_url_string,\
                equipments_name_string, equipments_power_string, equipments_kpp_string, equipments_fuel_string, equipment_old_price_string,\
                equipments_payment_per_month_string, equipments_benefit_string, equipments_special_credit_price_string,\
                equipments_sum_price_string, equipment_option_comfort_string, equipment_option_salon_string, equipment_option_visibility_string,\
                equipment_option_safety_string, equipment_option_theft_protection_string, equipment_option_multimedia_string,\
                equipment_option_exterior_elements_string, equipment_option_package_string, equipment_option_other_string,\
                equipments_2_options_string, engines_name_string, engines_kpp_string, engines_avto_class_string, engines_power_string, engines_capacity_string,\
                engines_drive_unit_string, engines_expenditure_string, engines_country_string, engines_engine_type_string, engines_fuel_string,\
                engines_amount_kpp_string]
    
    @staticmethod
    def get_engines_info(html_car_page):
        engines_name_string = str()
        engines_name_ntml = html_car_page.find('div', class_='sheet').find('div', class_='sheet__row').find_all('div', class_='sheet__col')[1::]
        for engine_name_ntml in engines_name_ntml: 
            engine_name = engine_name_ntml.text.replace('\n', '').lstrip().rstrip()
            engines_name_string += engine_name + '|'
        engines_name_string = engines_name_string[0:-1]
        
        engines_avto_class_string = str()
        engines_kpp_string = str()
        engines_power_string = str()
        engines_capacity_string = str()
        engines_drive_unit_string = str()
        engines_expenditure_string = str()
        engines_country_string = str()
        engines_engine_type_string = str()
        engines_fuel_string = str()
        engines_amount_kpp_string = str()
        engines_options_html = html_car_page.find('div', class_='sheet').find_all('div', class_='sheet__row')
        for option_html in engines_options_html:
            name_option = option_html.find('div').text.replace('\n', '').lstrip().rstrip()
            if name_option == 'Класс автомобиля':
                for avto_class in option_html.find_all('div')[1::]:
                    avto_class_text = avto_class.text.replace('\n', '').lstrip().rstrip()
                    if avto_class_text == '':
                        engines_avto_class_string += ' ' + '||'
                    else:
                        engines_avto_class_string += avto_class_text + '||'
            elif name_option == 'Коробка':
                for kpp in option_html.find_all('div')[1::]:
                    kpp_text = kpp.text.replace('\n', '').lstrip().rstrip()
                    if kpp_text == '':
                        engines_kpp_string += ' ' + '||'
                    else:
                        engines_kpp_string += kpp_text + '||'
            elif name_option == 'Мощность':
                for power in option_html.find_all('div')[1::]:
                    power_text = power.text.replace('\n', '').lstrip().rstrip()
                    if power_text == '':
                        engines_power_string += ' ' + '||'
                    else:
                        engines_power_string += power_text + '||'
            elif name_option == 'Объем':
                for capacity in option_html.find_all('div')[1::]:
                    capasity_text = capacity.text.replace('\n', '').lstrip().rstrip()
                    if capasity_text == '':
                        engines_capacity_string += ' ' + '||'
                    else:
                        engines_capacity_string += capasity_text + '||'
            elif name_option == 'Привод':
                for drive_unit in option_html.find_all('div')[1::]:
                    drive_unit_text = drive_unit.text.replace('\n', '').lstrip().rstrip()
                    if drive_unit_text == '':
                        engines_drive_unit_string += ' ' + '||'
                    else:
                        engines_drive_unit_string += drive_unit_text + '||'
            elif name_option == 'Расход':
                for expenditure in option_html.find_all('div')[1::]:
                    expenditure_text = expenditure.text.replace('\n', '').lstrip().rstrip()
                    if expenditure_text == '':
                        engines_expenditure_string += ' ' + '||'
                    else:
                        engines_expenditure_string += expenditure_text + '||'
            elif name_option == 'Страна марки':
                for country in option_html.find_all('div')[1::]:
                    country_text = country.text.replace('\n', '').lstrip().rstrip()
                    if country_text == '':
                        engines_country_string += ' ' + '||'
                    else:
                        engines_country_string += country_text + '||'
            elif name_option == 'Тип двигателя':
                for engine_type in option_html.find_all('div')[1::]:
                    engine_type_text = engine_type.text.replace('\n', '').lstrip().rstrip()
                    if engines_engine_type_string != '':
                        continue
                    if engine_type_text == '':
                        engines_engine_type_string += ' ' + '||'
                    else:
                        engines_engine_type_string += engine_type_text + '||'
            elif name_option == 'Топливо':
                for fuel in option_html.find_all('div')[1::]:
                    fuel_text = fuel.text.replace('\n', '').lstrip().rstrip()
                    if fuel_text == '':
                        engines_fuel_string += ' ' + '||'
                    else:
                        engines_fuel_string += fuel_text + '||'
            elif name_option == 'Количество передач':
                for amount_kpp in option_html.find_all('div')[1::]:
                    amount_kpp_text = amount_kpp.text.replace('\n', '').lstrip().rstrip()
                    if amount_kpp_text == '':
                        engines_amount_kpp_string += ' ' + '||'
                    else:
                        engines_amount_kpp_string += amount_kpp_text + '||'
            
        engines_kpp_string = engines_kpp_string[0:-2]
        engines_avto_class_string = engines_avto_class_string[0:-2]
        engines_power_string = engines_power_string[0:-2]
        engines_capacity_string = engines_capacity_string[0:-2]
        engines_drive_unit_string = engines_drive_unit_string[0:-2]
        engines_expenditure_string = engines_expenditure_string[0:-2]
        engines_country_string = engines_country_string[0:-2]
        engines_engine_type_string = engines_engine_type_string[0:-2]
        engines_fuel_string = engines_fuel_string[0:-2]
        engines_amount_kpp_string = engines_amount_kpp_string[0:-2]
        return [engines_name_string, engines_kpp_string, engines_avto_class_string, engines_power_string, engines_capacity_string, engines_drive_unit_string,\
                engines_expenditure_string, engines_country_string, engines_engine_type_string, engines_fuel_string, engines_amount_kpp_string]

    @staticmethod
    def get_equipmets_info(html_car_page):
        equipments_html = html_car_page.find('div', class_='table d-lg-block table-eq-desk')\
                                        .find_all('div', class_='col')[1:-1]
        equipments_name_string = str()
        equipments_power_string = str()
        equipments_kpp_string = str()
        equipments_fuel_string = str()
        equipment_old_price_string = str()
        equipments_payment_per_month_string = str()
        equipments_benefit_string = str()
        equipments_special_credit_price_string = str()
        equipments_sum_price_string = str()
        equipment_option_comfort_string = str()
        equipment_option_salon_string = str()
        equipment_option_visibility_string = str()
        equipment_option_safety_string = str()
        equipment_option_theft_protection_string = str()
        equipment_option_multimedia_string = str()
        equipment_option_exterior_elements_string = str()
        equipment_option_package_string = str()
        equipment_option_other_string = str()
        
        for equipment_html in equipments_html:
            equipment_fields_html = equipment_html.find_all('div', class_='table-body__item')
            equipment_name = equipment_fields_html[0].text.replace('\n', '').replace(' ', '')
            equipment_power = equipment_fields_html[1].text.replace('\n', '').replace(' ', '')
            equipment_kpp = equipment_fields_html[2].text.replace('.', '').replace('\n', '').replace(' ', '')
            equipment_fuel = equipment_fields_html[3].text.replace('\n', '').replace(' ', '')
            equipment_old_price = equipment_fields_html[4].find('span').text.replace(' ', '').replace('р', '')
            equipment_payment_per_month = equipment_fields_html[5].text.replace('\n', '').replace(' ', '').replace('от', '')\
                                                                        .replace('р**Всравнение', '').lstrip().rstrip()
            equipment_benefit = equipment_fields_html[6].text.replace('\n', '').replace(' ', '').replace('рЗарезервировать', '')
            equipment_special_credit_price = equipment_fields_html[7].text.replace('\n', '').replace(' ', '').replace('от', '')\
                                                                            .replace('рВкредит', '')
            equipment_sum_price = str(int(equipment_benefit) + int(equipment_special_credit_price))


            equipment_option_comfort = 'отсутствует'
            equipment_option_salon = 'отсутствует'
            equipment_option_visibility = 'отсутствует'
            equipment_option_safety = 'отсутствует'
            equipment_option_theft_protection = 'отсутствует'
            equipment_option_multimedia = 'отсутствует'
            equipment_option_exterior_elements = 'отсутствует'
            equipment_option_option_package = 'отсутствует'
            equipment_option_other = 'отсутствует'
            equipment_options_html = equipment_html.find('div', class_='row-opened__wrap').find_all('ul', class_='row-opened__list')

            for equipment_option_html in equipment_options_html:
                name_option = equipment_option_html.find('p').text
                if 'Комфорт' in name_option:
                    details_option = equipment_option_html.find_all('li')
                    for detail_option in details_option:
                        if detail_option.attrs['class'][0] == 'row-opened__list-item row-opened__list-item_choose':
                            equipment_option_comfort += detail_option.text + '|'
                        else: 
                            equipment_option_comfort += detail_option.find('span').text + '|'
                    equipment_option_comfort = equipment_option_comfort.replace('отсутствует', '')[0:-1]
                if 'Салон' in name_option:
                    details_option = equipment_option_html.find_all('li')
                    for detail_option in details_option:
                        if detail_option.attrs['class'][0] == 'row-opened__list-item row-opened__list-item_choose':
                            equipment_option_salon += detail_option.text + '|'
                        else: 
                            equipment_option_salon += detail_option.find('span').text + '|'
                    equipment_option_salon = equipment_option_salon.replace('отсутствует', '')[0:-1]
                if 'Обзор' in name_option:
                    details_option = equipment_option_html.find_all('li')
                    for detail_option in details_option:
                        if detail_option.attrs['class'][0] == 'row-opened__list-item row-opened__list-item_choose':
                            equipment_option_visibility += detail_option.text + '|'
                        else: 
                            equipment_option_visibility += detail_option.find('span').text + '|'
                    equipment_option_visibility = equipment_option_visibility.replace('отсутствует', '')[0:-1]
                if 'Безопасность' in name_option:
                    details_option = equipment_option_html.find_all('li')
                    for detail_option in details_option:
                        if detail_option.attrs['class'][0] == 'row-opened__list-item row-opened__list-item_choose':
                            equipment_option_safety += detail_option.text + '|'
                        else: 
                            equipment_option_safety += detail_option.find('span').text + '|'
                    equipment_option_safety = equipment_option_safety.replace('отсутствует', '')[0:-1]
                if 'Защита от угона' in name_option:
                    details_option = equipment_option_html.find_all('li')
                    for detail_option in details_option:
                        if detail_option.attrs['class'][0] == 'row-opened__list-item row-opened__list-item_choose':
                            equipment_option_theft_protection += detail_option.text + '|'
                        else: 
                            equipment_option_theft_protection += detail_option.find('span').text + '|'
                    equipment_option_theft_protection = equipment_option_theft_protection.replace('отсутствует', '')[0:-1]
                if 'Мультимедиа' in name_option:
                    details_option = equipment_option_html.find_all('li')
                    for detail_option in details_option:
                        if detail_option.attrs['class'][0] == 'row-opened__list-item row-opened__list-item_choose':
                            equipment_option_multimedia += detail_option.text + '|'
                        else: 
                            equipment_option_multimedia += detail_option.find('span').text + '|'
                    equipment_option_multimedia = equipment_option_multimedia.replace('отсутствует', '')[0:-1]
                if 'Элементы экстерьера' in name_option:
                    details_option = equipment_option_html.find_all('li')
                    for detail_option in details_option:
                        if detail_option.attrs['class'][0] == 'row-opened__list-item row-opened__list-item_choose':
                            equipment_option_exterior_elements += detail_option.find('i').text + '|'
                        else: 
                            if detail_option.find('i') is not None:
                                equipment_option_exterior_elements += detail_option.find('i').text + '|'
                            else:
                                equipment_option_exterior_elements += detail_option.find('span').text + '|'
                    equipment_option_exterior_elements = equipment_option_exterior_elements.replace('отсутствует', '')[0:-1]
                if 'Пакеты опций' in name_option:
                    details_option = equipment_option_html.find_all('li')
                    for detail_option in details_option:
                        if detail_option.attrs.get('class') is not None:
                            if detail_option.attrs['class'][0] == 'row-opened__list-item row-opened__list-item_choose':
                                equipment_option_option_package += detail_option.text + '|'
                            else: 
                                if detail_option.find('label') is not None:
                                    equipment_option_option_package += detail_option.find('label').find('i').text + '|'
                                else:
                                    equipment_option_option_package += detail_option.find('span').text + '|'
                    equipment_option_option_package = equipment_option_option_package.replace('отсутствует', '')[0:-1]

                if 'Прочее' in name_option:
                    details_option = equipment_option_html.find_all('li')
                    for detail_option in details_option:
                        if detail_option.attrs['class'][0] == 'row-opened__list-item row-opened__list-item_choose':
                           equipment_option_other += detail_option.text + '|'
                        else: 
                           equipment_option_other += detail_option.find('span').text + '|'
                    equipment_option_other = equipment_option_other.replace('отсутствует', '')[0:-1]

            equipment_option_comfort_string += equipment_option_comfort + '||'
            equipment_option_salon_string += equipment_option_salon + '||'
            equipment_option_visibility_string += equipment_option_visibility + '||'
            equipment_option_safety_string += equipment_option_safety + '||'
            equipment_option_theft_protection_string += equipment_option_theft_protection + '||'
            equipment_option_multimedia_string += equipment_option_multimedia + '||'
            equipment_option_exterior_elements_string += equipment_option_exterior_elements + '||'
            equipment_option_package_string += equipment_option_option_package + '||'
            equipment_option_other_string += equipment_option_other + '||'

            equipments_name_string += equipment_name + '||'
            equipments_power_string += equipment_power + '||'
            equipments_kpp_string += equipment_kpp + '||'
            equipments_fuel_string += equipment_fuel + '||'
            equipment_old_price_string += equipment_old_price + '||'
            equipments_payment_per_month_string += equipment_payment_per_month + '||'
            equipments_benefit_string += equipment_benefit + '||'
            equipments_special_credit_price_string += equipment_special_credit_price + '||'
            equipments_sum_price_string += equipment_sum_price + '||'

        return [equipments_name_string[0:-2], equipments_power_string[0:-2], equipments_kpp_string[0:-2], equipments_fuel_string[0:-2],\
                equipment_old_price_string[0:-2], equipments_payment_per_month_string[0:-2], equipments_benefit_string[0:-2],\
                equipments_special_credit_price_string[0:-2], equipments_sum_price_string[0:-2], equipment_option_comfort_string[0:-2],\
                equipment_option_salon_string[0:-2], equipment_option_visibility_string[0:-2], equipment_option_safety_string[0:-2],\
                equipment_option_theft_protection_string, equipment_option_multimedia_string[0:-2],\
                equipment_option_exterior_elements_string[0:-2], equipment_option_package_string[0:-2], equipment_option_other_string[0:-2]]

    
    def get_color_info(self, html_car_page):
        all_color_html = html_car_page.find('ul', class_='offer-main__color').find_all('li')
        color_name_string = str()
        color_code_string = str()
        color_url_string = str()
        for color_html in all_color_html:
            color_name = color_html.get('data-name')
            color_code = color_html.get('style').split(' ')[-1].replace(';', '')
            color_url = self.MAIN_URL + color_html.get('data-img')
            if color_name == '':
                color_name = 'нет названия'
            if color_code == '':
                color_code = 'нет кода'
            color_name_string += color_name + '|'
            color_code_string += color_code + '|'
            color_url_string += color_url + '|'
        return color_name_string[0:-1], color_code_string[0:-1], color_url_string[0:-1]
    
    def write_cars_info(self, info_list):
        book = xlsxwriter.Workbook(f'автомобили.xlsx')
        page = book.add_worksheet('автомобили')

        row = 0
        column = 0

        page.set_column('A:A', 50)
        page.set_column('B:B', 50)
        page.set_column('C:C', 50)
        page.set_column('D:D', 50)
        page.set_column('E:E', 50)
        page.set_column('F:F', 50)
        page.set_column('G:G', 50)
        page.set_column('H:H', 50)
        page.set_column('I:I', 50)
        page.set_column('J:J', 50)
        page.set_column('K:K', 50)
        page.set_column('L:L', 50)
        page.set_column('M:M', 50)
        page.set_column('N:N', 50)
        page.set_column('O:O', 50)
        page.set_column('P:P', 50)
        page.set_column('Q:Q', 50)
        page.set_column('R:R', 50)
        page.set_column('S:S', 50)
        page.set_column('T:T', 50)
        page.set_column('U:U', 50)
        page.set_column('V:V', 50)
        page.set_column('W:W', 50)
        page.set_column('X:X', 50)
        page.set_column('Y:Y', 50)
        page.set_column('Z:Z', 50)
        page.set_column('AA:AA', 50)
        page.set_column('AB:AB', 50)
        page.set_column('AC:AC', 50)
        page.set_column('AD:AD', 50)
        page.set_column('AE:AE', 50)
        page.set_column('AF:AF', 50)
        page.set_column('AG:AG', 50)
        page.set_column('AH:AH', 50)
        page.set_column('AI:AI', 50)
        page.set_column('AJ:AJ', 50)
        page.set_column('AK:AK', 50)
        page.set_column('AL:AL', 50)
        page.set_column('AM:AM', 50)
        page.set_column('AN:AN', 50)
        page.set_column('AO:AO', 50)

        page.write(row, column, 'бренд')
        page.write(row, column+1, 'модель')
        page.write(row, column+2, 'полное название')
        page.write(row, column+3, 'главное фото')
        page.write(row, column+4, 'актуальная цена модели')
        page.write(row, column+5, 'старая цена модели')
        page.write(row, column+6, 'мощьность модели')
        page.write(row, column+7, 'объем модели')
        page.write(row, column+8, 'названия цветов')
        page.write(row, column+9, 'коды цветов')
        page.write(row, column+10, 'ссылки цветов')
        page.write(row, column+11, 'названия комплектаций')
        page.write(row, column+12, 'мощьность компелктаций')
        page.write(row, column+13, 'кпп комплектаций')
        page.write(row, column+14, 'топливо комплектаций')
        page.write(row, column+15, 'старая цена комплектации')
        page.write(row, column+16, 'оплата в месяц')
        page.write(row, column+17, 'выгода')
        page.write(row, column+18, 'спец. кредитная цена')
        page.write(row, column+19, 'сумма')
        page.write(row, column+20, 'комфорт')
        page.write(row, column+21, 'салон')
        page.write(row, column+22, 'обзор')
        page.write(row, column+23, 'безопасность')
        page.write(row, column+24, 'защита от угона')
        page.write(row, column+25, 'мультимедия')
        page.write(row, column+26, 'элементы экстерьера')
        page.write(row, column+27, 'пакет опций')
        page.write(row, column+28, 'другое')
        page.write(row, column+29, 'комплектация 2')
        page.write(row, column+30, 'название двигателя')
        page.write(row, column+31, 'двиг. коробка передач')
        page.write(row, column+32, 'двиг. класс авто')
        page.write(row, column+33, 'двиг. мощность')
        page.write(row, column+34, 'двиг. объем')
        page.write(row, column+35, 'двиг. привод')
        page.write(row, column+36, 'двиг. расход')
        page.write(row, column+37, 'двиг. страна')
        page.write(row, column+38, 'двиг. тип двигателя')
        page.write(row, column+39, 'двиг. топливо')
        page.write(row, column+40, 'двиг. число передач')

        row += 1

        for item_info in info_list:
            page.write(row, column, item_info[0])
            page.write(row, column+1, item_info[1])
            page.write(row, column+2, item_info[2])
            page.write(row, column+3, item_info[3])
            page.write(row, column+4, item_info[4])
            page.write(row, column+5, item_info[5])
            page.write(row, column+6, item_info[6])
            page.write(row, column+7, item_info[7])
            page.write(row, column+8, item_info[8])
            page.write(row, column+9, item_info[9])
            page.write(row, column+10, item_info[10])
            page.write(row, column+11, item_info[11])
            page.write(row, column+12, item_info[12])
            page.write(row, column+13, item_info[13])
            page.write(row, column+14, item_info[14])
            page.write(row, column+15, item_info[15])
            page.write(row, column+16, item_info[16])
            page.write(row, column+17, item_info[17])
            page.write(row, column+18, item_info[18])
            page.write(row, column+19, item_info[19])
            page.write(row, column+20, item_info[20])
            page.write(row, column+21, item_info[21])
            page.write(row, column+22, item_info[22])
            page.write(row, column+23, item_info[23])
            page.write(row, column+24, item_info[24])
            page.write(row, column+25, item_info[25])
            page.write(row, column+26, item_info[26])
            page.write(row, column+27, item_info[27])
            page.write(row, column+28, item_info[28])
            page.write(row, column+29, item_info[29])
            page.write(row, column+30, item_info[30])
            page.write(row, column+31, item_info[31])
            page.write(row, column+32, item_info[32])
            page.write(row, column+33, item_info[33])
            page.write(row, column+34, item_info[34])
            page.write(row, column+35, item_info[35])
            page.write(row, column+36, item_info[36])
            page.write(row, column+37, item_info[37])
            page.write(row, column+38, item_info[38])
            page.write(row, column+39, item_info[39])
            page.write(row, column+40, item_info[40])

            row += 1


        book.close()
    

if __name__ == '__main__':
    parser = Parser()
    brand_list = parser.get_all_url_brand()
    cars_url_list = parser.pars_brand_list(brand_list)
    cars_info_list = parser.pars_cars_list(cars_url_list)
    parser.write_cars_info(cars_info_list)