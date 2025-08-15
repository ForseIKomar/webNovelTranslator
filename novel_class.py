class BookParser:
    def __init__(self, res_arr, book):
        self.info_str = res_arr[0][0]
        self.book = book

    def parse(self):
        print(self.info_str)
        if 'panda-novel' in self.info_str:
            self._parse_panda_novel()
        elif 'fanfiction.net' in self.info_str:
            self._parse_fanfiction_net()
        elif 'hostednovel.com' in self.info_str:
            self._parse_hostednovel()
        elif 'Custom-1' in self.info_str:
            self._parse_custom_1()
        elif ' genesistls.com' in self.info_str:
            self._parse_genesis_tls()
        elif 'www.isotls.com' in self.info_str:
            self._parse_isotls()
        elif 'noblemtl.com' in self.info_str:
            self._parse_noblemtl()
        elif 'webnovel.com' in self.info_str:
            self._parse_webnovel()
        elif 'scribblehub.com' in self.info_str:
            self._parse_scribblehub()
        elif 'Следов:' in self.info_str:
            self._parse_sledov()
        else:
            self._parse_default()

        self.book.set_language('ru')
        return self.book

    def _parse_panda_novel(self):
        self.book.set_identifier(self.info_str.split('|')[2].split('/')[-1])
        self.book.set_title(self.info_str.split('|')[3])
        self.book.add_author(self.info_str.split('|')[9])

    def _parse_fanfiction_net(self):
        self.book.set_identifier(self.info_str.split('|')[3].split(': ')[-1])
        element = ''
        for elem in self.info_str.split('|'):
            if 'fanfiction.net' in elem:
                element = elem
        self.book.set_title(element.split('/')[-1])
        self.book.add_author(self.info_str.split('|')[6])

    def _parse_hostednovel(self):
        self.book.set_identifier(self.info_str.split('|')[2])
        self.book.set_title(self.info_str.split('|')[2].split('/')[-1].split('#')[0])

    def _parse_custom_1(self):
        self.book.set_identifier(self.info_str.split('|')[0])
        self.book.set_title(self.info_str.split('|')[0])

    def _parse_genesis_tls(self):
        self.book.set_identifier(self.info_str.split('|')[2])
        self.book.set_title(self.info_str.split('|')[2].split('/')[-2])

    def _parse_isotls(self):
        self.book.set_identifier(self.info_str.split('|')[2])
        self.book.set_title(self.info_str.split('|')[2].split('/')[-2])

    def _parse_noblemtl(self):
        self.book.set_identifier(self.info_str.split('|')[2])
        self.book.set_title(self.info_str.split('|')[2].split('/')[-2])

    def _parse_webnovel(self):
        self.book.set_identifier(self.info_str.split('|')[2])
        self.book.set_title(self.info_str.split('|')[3])

    def _parse_scribblehub(self):
        self.book.set_identifier(self.info_str.split('|')[2])
        self.book.set_title(self.info_str.split('|')[2].split('/')[-2])

    def _parse_sledov(self):
        self.book.set_identifier(self.info_str.split('|')[0].split(': ')[-1])
        self.book.set_title(self.info_str.split('|')[0].split('/')[-1])
        self.book.add_author(self.info_str.split('|')[4])

    def _parse_default(self):
        self.book.set_identifier(self.info_str.split('|')[2])
        self.book.set_title(self.info_str.split('|')[2])





