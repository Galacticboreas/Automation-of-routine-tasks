class ArticleExtractor:

    articles = dict()
    num_of_iter = 0

    def get_article(self, name):
        self.num_of_iter += 1
        if self.articles.get(name):
            return self.articles[name]
        return self.__extract_atricle(name)

    def __extract_atricle(self, name):
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
        num_right = name.rfind(")")
        num_left = name.rfind("(") + 1
        return num_right - num_left

    def get_symbol(self, name):
        return name[name.rfind("(") + 1: name.rfind(")")]

    def get_num_extract_articles(self):
        return len(self.articles)

    def get_num_of_iter(self):
        return self.num_of_iter


def test_func_articles(a, b):
    return f'a = {a}, b = {b}'
