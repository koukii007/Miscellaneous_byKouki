import scrapy
from datetime import date
from datetime import timedelta
from scrapy.crawler import CrawlerProcess



class Top100Spider(scrapy.Spider):

    name = 'top100'
    allowed_domains = ['www.billboard.com']
    start_urls = ['https://www.billboard.com/charts/artist-100/']

    def __init__(self):

        self.today = date.today()
        self.list_of_artists = []
    def parse(self, response):
        extract = response.css('.a-no-trucate::text').getall()
        for element in range(0,5):

            artist_name = extract[element].replace("\n", '')
            artist_name1 = artist_name.replace("\t", '')

            if artist_name1 not in self.list_of_artists:
                self.list_of_artists.append(artist_name1)
                yield {'Artist': artist_name1}
        for i in range(1, 4):
            previous_week = self.today - timedelta(days=i * 7)
            next_page = 'https://www.billboard.com/charts/artist-100/' + str(previous_week)
            if next_page:
                yield response.follow(next_page, callback=self.parse)

    def close(self, spider, reason):
        with open(settings.FEED_URI, 'r+') as myfile:
            load_lines = myfile.readlines()
            load_lines.pop(0)
            load_lines = sorted(load_lines)
            load_lines.insert(0,'Artists\n')
            myfile.seek(0)
            for item in load_lines:
                myfile.write(item)
            myfile.close()

//
process = CrawlerProcess(settings={
    "FEEDS": {
        "Artists.csv": {"format": "csv"},
    },
})

process.crawl(Top100Spider)
process.start()