class ArticleExtractor:
    """
    Клас предназначен для извлечения и хеширования 6-ти значного
    артикула из полного нименования изделия мебели.

    Пример №1: Тумба НекоеНазвание 60 белая (194236)
    Результат: 194236

    Пример №2: Тумба НекоеНазвание 60 белая (005) (194236)
    Результат: 194236

    Пример №3: Тумба НекоеНазвание (глянец) 60 белая (005) (194236)
    Результат: 194236

    Пример №4: Тумба НекоеНазвание (глянец) 60 белая (005) (194236).001
    Результат: 194236

    Returns:
        articles: [6-ти значные цифры арткула мебели]
    """

    articles = dict()
    num_of_iter = 0

    def get_article(self, name):
        """Функция возрвращает артикул если он есть в кэше,
        если нет, то вызывает функцию извлечения артикула

        Args:
            name (str): [полное наименование изделия мебели]

        Returns:
            article (int): [6-ти значный артикул изделия]
        """
        self.num_of_iter += 1
        if self.articles.get(name):
            return self.articles[name]
        return self.__extract_atricle(name)

    def __extract_atricle(self, name):
        """Функция извлекает 6-ти значный артикул из 
        полного наименования

        Args:
            name (str): [полное наименование изделия мебели]

        Returns:
            atricle (str): [6-ти значный артикул изделия]
        """
        if name:
            num_of_occur = name.count(")")
            temp_name = name
            for i in range(num_of_occur):
                if self.get_num_char(temp_name) == 6:
                    art_temp = self.get_symbol(temp_name)
                    if art_temp.isdigit():
                        self.articles[temp_name] = art_temp
                        return art_temp
                    else:
                        num_right = temp_name.rfind(")")
                        num_left = temp_name.rfind("(")
                        temp_name = temp_name[:num_right] + temp_name[num_right + 1:]
                        temp_name = temp_name[:num_left] + temp_name[num_left + 1:]
                else:
                    num_right = temp_name.rfind(")")
                    num_left = temp_name.rfind("(")
                    temp_name = temp_name[:num_right] + temp_name[num_right + 1:]
                    temp_name = temp_name[:num_left] + temp_name[num_left + 1:]
            self.articles[name] = name
            return name
        return name

    def get_num_char(self, name):
        """Функция определяет количество символов
        между скобками правой части строки.

        Args:
            name (str): строка со скобками

        Returns:
            int: количество символов между скобками (..кол-во..)
        """
        num_right = name.rfind(")")
        num_left = name.rfind("(") + 1
        return num_right - num_left

    def get_symbol(self, name):
        """Функция удаляет скобки из правой части строки

        Args:
            name (str): [строка со кобками]

        Returns:
            name (str): [строка с удаленными скобками первого вхождения
            правой части]
        """
        return name[name.rfind("(") + 1: name.rfind(")")]

    def get_num_extract_articles(self):
        """Подчсет количества извлеченных артикулов в кэше

        Returns:
            len(articles) (int): [длина словаря с артикулами]
        """
        return len(self.articles)

    def get_num_of_iter(self):
        """Подстчет количества итераций по строкам отчета

        Returns:
            num_of_iter (int): [количество итераций]
        """
        return self.num_of_iter
