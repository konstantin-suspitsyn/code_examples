class masterData:
    '''
    Файл предлагает категорию товарп по названию
    Файл проверяет по расстоянию Левинштейна потенциально заведенные товары в систему
    Файл проверяет размеры внутри группы (+/- 3σ)
    '''

    # Папка с импортом Excel
    IMPORT_FOLDER = r'import'
    # Папка с экспортом Excel
    EXPORT_FOLDER = r'export'
    # Файл с мастер-данными по SKU
    MD_FROM_DB = r'b-up/source.csv'
    #Папка с файлами для класса
    BUP = r'class_data'
    # Файл с мастер-данными по SKU
    MD_IMPORT = BUP + r'/sku_library.json'
    # Список удалаяемых символов без пробела
    DELETE_NO_SPACE = BUP + r'/no_space.yaml'
    # Список удалаяемых символов без пробела
    DELETE_BRANDS = BUP + r'/brands.yaml'
    # Список удалаяемых символов с пробелом
    DELETE_WITH_SPACE = BUP + r'/with_space.yaml'
    # Названия листа в файлах вендора
    EXCEL_VENDOR_SHEET = ['Товар в одной упаковке', 'Товар в нескольких упаковках']
    # Названия колонок от вендора
    EXCEL_COLUMNS = BUP + r'/columns.yaml'
    # Словарь для очистки от русских букв и нескольких других символов
    RUS_LETTERS = ['й', 'ц', 'у', 'к', 'е', 'н', 'г', 'ш', 'щ',
                   'з', 'х', 'ъ', 'ф', 'ы', 'в', 'а', 'п', 'р',
                   'о', 'л', 'д', 'ж', 'э', 'я', 'ч', 'с', 'м',
                   'и', 'т', 'т', 'ь', 'б', 'ю', ' ', '-', ')',
                   '(', '//', ';', ':', '/', '#', '№', '&']

    import pandas as pd
    import numpy as np
    import yaml
    import json
    import tqdm
    import glob

    def __init__(self):
        pass

    def list_files_in_custom_dir(self, create_folder=False, folder_name='Output', path=None):
        '''
        Получаем список файлов по пути, который должен вставить человек
        :param create_folder: если нужно создать папку пишем True
        :param folder_name: если create_folder=True, то название папки
        :return: список файлов. [0] - последняя правая папка, в которой лежат данные, [1] - список всех файлов с прямыми
        ссылками
        '''
        import os
        if path==None:
            print('Вставьте полный путь до папки:')
            path = input()
        files = [f for f in self.glob.glob(path + "*.xls*", recursive=True)]
        res = []
        # folder = files[0].split('\\')[len(files[0].split('\\'))-2]
        folder = path + folder_name + '\\'
        if create_folder==True:
            if not os.path.exists(folder):
                os.makedirs(folder)
        res.append(path)
        res.append(folder)
        res.append(files)
        return res

    def import_dict_from_json(self, link):
        f = open(link)
        res = self.json.load(f)
        f.close()
        return res

    def import_from_yaml(self, file, code='utf-8'):
        '''
        Импорт данных из yaml
        :param file: ссылка на yaml - файл
        :param code: кодировка файла. Лучше в utf-8
        :return: возвращает list
        '''
        with open(file, 'r', encoding=code) as file:
            data = self.yaml.safe_load(file)
        return data

    def remove_from_string(self, string):
        '''
        Очищает строку от определенных символов.
        Может просто удалять
        Может заменять на пробел
        :param string: строка с которой работаем
        symbols_list_no_space: символы, которые удаляются без пробела
        symbols_list_with_space: символы, которые удаляются и вместо них ставится пробел
        :return: очищенная строка string
        '''
        symbols_list_no_space = self.import_from_yaml(self.DELETE_NO_SPACE)
        symbols_list_with_space = self.import_from_yaml(self.DELETE_WITH_SPACE)
        result_string = str(string).lower()
        for s in symbols_list_no_space:
            result_string = result_string.replace(s.lower(), '')
        for s in symbols_list_with_space:
            result_string = result_string.replace(s.lower(), ' ')
        return result_string

    def split_SKU_no_numbers(self, string):
        '''
        Разделяем по пробелам название SKU
        Использует функцию remove_from_string
        :param string: строка, которую будем разбивать
        :return: list
        '''
        string = str(string).lower()
        string_list_names = self.remove_from_string(string).split(' ')
        string_no_numbers = []
        # Проверяем есть ли цифры в слове
        # Если есть, то решаем, что это не название, а модель
        for word in string_list_names:
            numbers = sum(n.isdigit() for n in word)
            if numbers == 0:
                # Если часть слова меньше 3 символов, исключаем
                if len(word) >= 3:
                    string_no_numbers.append(word)
        return string_no_numbers

    def avg_group_volume(self, base_hierarchy):
        '''
        Получаем средний размер по группе
        И отклонение +/- 3σ
        :param base_hierarchy: иерархия с SKU-группа-размеры
        :return: новый DataFrame c посчитанными отклоненениями от среднего размера
        '''
        avg_groupvolume = base_hierarchy
        avg_groupvolume['vol'] = base_hierarchy['length'] * base_hierarchy['height'] * base_hierarchy['width'] / (
                    100 ^ 3)
        avg_groupvolume = base_hierarchy[['group_name', 'vol']].groupby('group_name').agg(
            [self.np.std, 'mean']).reset_index()
        avg_groupvolume.columns = ['group_name', 'std_deviation', 'avg_group_volume']
        avg_groupvolume['avg_group_volume']=avg_groupvolume['avg_group_volume'].fillna(0)
        avg_groupvolume['std_deviation'] = avg_groupvolume['std_deviation'].fillna(0)
        avg_groupvolume['max_deviation'] = avg_groupvolume['avg_group_volume'] + 3 * avg_groupvolume['std_deviation']
        avg_groupvolume['min_deviation'] = avg_groupvolume['avg_group_volume'] - 3 * avg_groupvolume['std_deviation']
        return avg_groupvolume

    def distance_levenshtain(self, x, y, memo=None):
        '''
        Расстояние Левенштейна
        :param x: Слово 1
        :param y: Слово 2
        :param memo: нужно для рекурсии
        :return: Расстояние Левенштейна
        '''
        x = str(x)
        y = str(y)
        if memo is None: memo = {}
        if len(x) == 0: return len(y)
        if len(y) == 0: return len(x)
        if (len(x), len(y)) in memo:
            return memo[(len(x), len(y))]
        delt = 1 if x[-1] != y[-1] else 0
        diag = self.distance_levenshtain(x[:-1], y[:-1], memo) + delt
        vert = self.distance_levenshtain(x[:-1], y, memo) + 1
        horz = self.distance_levenshtain(x, y[:-1], memo) + 1
        ans = min(diag, vert, horz)
        memo[(len(x), len(y))] = ans
        return ans

    def clear_model(self, df):
        '''
        Поскольку, у нас нет чистых моделей, просто очистим от русских букв
        Функция для лямбды
        :param df: df базовой иерархии
        :return: очищенные наименования
        '''
        model = df['model'].lower()
        model = model.replace(df['brand'], '')
        for l in self.RUS_LETTERS:
            model = model.replace(l, '')
        return model

    def clear_model_vendor(self, df):
        '''
        Очищаем модели производителя от русских символов
        :param df: df c моделями производителя
        :return:
        '''
        model = str(df['Артикул производителя']).lower()
        for l in self.RUS_LETTERS:
            model = model.replace(l, '')
        return model

    def create_master_data(self):
        '''
        Пересоздание файла-библиотеки
        для корректного определения группы
        :return: json-файл библиотеки
        '''

        base_hierarchy=self.pd.read_csv(self.MD_FROM_DB)
        exclude_from_hierarchy = base_hierarchy['brand'].str.lower().unique()

        with open(self.DELETE_BRANDS, 'w') as outfile:
            self.yaml.dump(exclude_from_hierarchy.tolist(), outfile, default_flow_style=False)

        base_hierarchy=base_hierarchy[['product_name', 'group_name']]

        group_hierarchy_dict = {}

        print('Начинаю создание словаря')
        #Создание словаря
        for index, row in self.tqdm.tqdm(base_hierarchy.iterrows(), total=base_hierarchy.shape[0]):
            SKU_name = row['product_name']
            # В каждое слово из разбивки названия, добавляем частоту встречающихся вещей
            for SKU_part in self.split_SKU_no_numbers(SKU_name):
                # Если слово - это бренд товара, то исключаем
                if SKU_part not in exclude_from_hierarchy:
                    if SKU_part in group_hierarchy_dict:
                        if row['group_name'] in group_hierarchy_dict[SKU_part]:
                            group_hierarchy_dict[SKU_part][row['group_name']] += 1
                        else:
                            # Все МП категории получают наивысший приоритет
                            if row['group_name'][:3] == 'МП_':
                                group_hierarchy_dict[SKU_part][row['group_name']] = 9999
                            else:
                                group_hierarchy_dict[SKU_part][row['group_name']] = 1
                    else:
                        if row['group_name'][:3] == 'МП_':
                            group_hierarchy_dict[SKU_part] = {row['group_name']: 9999}
                        else:
                            group_hierarchy_dict[SKU_part] = {row['group_name']: 1}

        json = self.json.dumps(group_hierarchy_dict)
        f = open(self.MD_IMPORT, "w+")
        f.write(json)
        f.close()

        print('Сохраняю результат в файл')

        del SKU_name
        del base_hierarchy
        del index
        del row
        del json
        print('Пересоздание мастер-библиотеки выполнено')

    def set_groups_xl(self):
        '''
        Проставляет группы к названиям
        :return: Excel-файлы
        '''
        import functools
        import operator
        import collections
        # Импорт словаря из сохраненного в json
        # Словарь пересоздается функцией create_master_data(). Должен запускаться отдельно
        group_hierarchy_dict = self.import_dict_from_json(self.MD_IMPORT)
        # получаем список файлов из папки. Она приходит на вход от пользователя
        filelist = self.list_files_in_custom_dir(create_folder=True, folder_name='Groups')
        # Перебираем все файлы в списке
        i = 1
        for link in filelist[2]:
            file_name = link.split('\\')[len(link.split('\\'))-1]
            print('Обрабатываем: {} из {}'.format(i, len(filelist[2])))
            print(file_name)

            # Определяем какие листы есть в Excel есть, а каких нет из списка
            xl = self.pd.ExcelFile(link).sheet_names
            xl_sheets=set(xl) & set(self.EXCEL_VENDOR_SHEET)
            for sheet in xl_sheets:
                from_vendor = self.pd.read_excel(link, sheet_name=sheet)
                from_vendor_SKU = from_vendor[['Наименование товара в системе поставщика (в отгрузочных документах ТОРГ-12 или УПД)']]
                if from_vendor_SKU.iloc[:,0].count() > 0:
                    vendor_SKUs_category = []
                    # Присваиваем словари - классификатор
                    for index, row in self.tqdm.tqdm(from_vendor_SKU.iterrows(), total=from_vendor_SKU.shape[0]):
                        SKU_name = str(row['Наименование товара в системе поставщика (в отгрузочных документах ТОРГ-12 или УПД)']).lower()
                        SKU_category = []
                        SKU_category.append(SKU_name)
                        dicts = []
                        # В каждое слово из разбивки названия, добавляем словарь категорий
                        for SKU_part in self.split_SKU_no_numbers(SKU_name):
                            if SKU_part in group_hierarchy_dict:
                                dicts.append(group_hierarchy_dict[SKU_part])
                        if dicts == []:
                            dicts = [{'None': 1}]
                        result_dict = dict(functools.reduce(operator.add, map(collections.Counter, dicts)))
                        result_dict = {k: v for k, v in sorted(result_dict.items(), key=lambda item: item[1], reverse=True)}
                        for k in list(result_dict.keys())[:5]:
                            SKU_category.append(k)
                        vendor_SKUs_category.append(SKU_category)

                    categoryDF = self.pd.DataFrame(vendor_SKUs_category,
                                              columns=['Название', 'Top_1', 'Top_2', 'Top_3', 'Top_4', 'Top_5'])
                    categoryDF['File_name']=file_name
                    categoryDF['sheet']=sheet
                    categoryDF = categoryDF[categoryDF['Название'].notna()]

                    # Добавляем нужные колонки в файл, чтобы проверить размеры и модель
                    # Обсудить финальный список колонок
                    columns_t=self.import_from_yaml(self.EXCEL_COLUMNS)['base_excel_columns']
                    from_vendor=from_vendor[columns_t]
                    col_to_add=self.import_from_yaml(self.EXCEL_COLUMNS)['columns_to_add'][0]
                    columns_t.append(col_to_add)
                    from_vendor[col_to_add]=from_vendor['Наименование товара в системе поставщика (в отгрузочных документах ТОРГ-12 или УПД)'].str.lower()
                    categoryDF=categoryDF.merge(from_vendor[columns_t],
                                                on='Название', how='left')
                    col_to_add=self.import_from_yaml(self.EXCEL_COLUMNS)['columns_to_add'][1]
                    list_sizes=self.import_from_yaml(self.EXCEL_COLUMNS)['sizes']
                    categoryDF[col_to_add]=categoryDF[list_sizes[0]] * \
                                           categoryDF[list_sizes[1]] * \
                                           categoryDF[list_sizes[2]] / (100**3)
                    try:
                        resDF = resDF.append(categoryDF)
                    except:
                        resDF = categoryDF
            resDF=resDF[resDF['Название']!='nan']
            resDF.to_excel(filelist[1]+file_name.lower(), sheet_name='Data', index=False)
            del resDF
            i+=1
        print('Группы предположены. Файлы сохранены. Проверяйте')

    def check_sizes_and_similarity(self):
        # получаем список файлов из папки. Она приходит на вход от пользователя
        filelist = self.list_files_in_custom_dir(create_folder=True, folder_name='Output')
        # просто создаем еще одну папку
        filelist = self.list_files_in_custom_dir(create_folder=True, folder_name='Levinstein', path=filelist[0])
        # Текущие мастер-данные по то товарам
        df_base_hierarchy = self.pd.read_csv(self.MD_FROM_DB)
        # Получаем ожидаемые размеры от товара внутри группы
        df_sizes_a = self.avg_group_volume(df_base_hierarchy)
        # Чистим от русских букв и пробелов модели
        df_base_hierarchy['model_cleared'] = df_base_hierarchy.apply(lambda x: self.clear_model(x), axis=1)
        df_base_hierarchy['model_cleared'] = df_base_hierarchy['model_cleared'].astype('str')
        # Список колонок, которые мы забираем из файлов шага set_groups_xl
        columns = self.import_from_yaml(self.EXCEL_COLUMNS)['base_excel_columns']
        added_columns=self.import_from_yaml(self.EXCEL_COLUMNS)['columns_to_add']
        for col in added_columns:
            columns.append(col)

        columns.append('Top_1')

        vol_col = self.import_from_yaml(self.EXCEL_COLUMNS)['columns_to_add'][1]

        # Обрабатываем все файлы в папке
        i=1
        for f in filelist[2]:
            file_name = f.split('\\')[len(f.split('\\')) - 1]
            print('Обрабатываем файл {} из {}'.format(i, len(filelist[2])))
            print(file_name)
            i+=1
            df = self.pd.read_excel(filelist[0] + 'Groups\\' + f.split('\\')[len(f.split('\\')) - 1])
            df = df[columns]
            try:
                df_groups.append(df)
            except:
                df_groups = df
            df_groups = df_groups.merge(df_sizes_a, left_on='Top_1', right_on='group_name', how='left')
            df_groups['vol_ok'] = df_groups.apply(lambda x: 'Ok' if (
                        (x[vol_col] <= x['max_deviation']) &
                        (x[vol_col] >= x[
                            'min_deviation'])) else 'Not Ok', axis=1)
            df_groups['Артикул для Левинштейна'] = df_groups.apply(lambda x: self.clear_model_vendor(x), axis=1)
            df_levi = df_groups[['Наименование товара в системе поставщика (в отгрузочных документах ТОРГ-12 или УПД)',
                                 'Артикул производителя',
                                 'Артикул для Левинштейна',
                                 'group_name']].merge(
                df_base_hierarchy[['group_name', 'product_name', 'model_cleared']],
                on='group_name', how='left')
            df_levi['lev_dist'] = df_levi.apply(
                lambda x: self.distance_levenshtain(x['Артикул для Левинштейна'], x['model_cleared']), axis=1)
            df_levi = df_levi.sort_values(by=['Артикул производителя', 'lev_dist'])
            df_levi = df_levi[df_levi['lev_dist']<=2]
            df_levi['Filter'] = None
            df_groups.to_excel(filelist[0]+'Output\\'+file_name, index=False)
            df_levi.to_excel(filelist[0]+'Levinstein\\'+file_name, index=False)
            del df_groups
        print('Готово')

    def final_correction(self):
        filelist = self.list_files_in_custom_dir(create_folder=False)
        i=0
        for f in filelist[2]:
            print('Часть 1 из 2')
            file_name = f.split('\\')[len(f.split('\\')) - 1]
            print('Обрабатываем файл {} из {}'.format(i, len(filelist[2])))
            print(file_name)
            df = self.pd.read_excel(filelist[0] + 'Levinstein\\' + f.split('\\')[len(f.split('\\')) - 1])
            try:
                df_levi = df_levi.append(df)
            except:
                df_levi = df

        list_levi = df_levi['Наименование товара в системе поставщика (в отгрузочных документах ТОРГ-12 или УПД)'][df_levi['Filter'].notnull()].unique().tolist()
        del df_levi
        i=0
        for f in filelist[2]:
            print('Часть 2 из 2')
            file_name = f.split('\\')[len(f.split('\\')) - 1]
            print('Обрабатываем файл {} из {}'.format(i, len(filelist[2])))
            print(file_name)
            file_name = f.split('\\')[len(f.split('\\')) - 1]
            df = self.pd.read_excel(filelist[0] + 'Output\\' + f.split('\\')[len(f.split('\\')) - 1])
            df = df[~df['Наименование товара в системе поставщика (в отгрузочных документах ТОРГ-12 или УПД)'].isin(list_levi)]
            df.to_excel(filelist[0]+'Output\\'+file_name, index=False)
        print('Готово')

