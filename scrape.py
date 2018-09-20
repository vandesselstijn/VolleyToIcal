from lxml import html
import requests


print('Conecting to https://wavoc.be/kalender/\n')
page = requests.get('https://wavoc.be/kalender/')

print('Scraping matches\n')
tree = html.fromstring(page.content)

matches = tree.xpath('//td/text()|//strong/text()')

print('Matches: ', matches)

print(matches[0], matches[3], matches[5])

print('\n')

#for i in range(len(matches))
    #print(matches[0], matches[3], matches[5])


#ploegen
#ploegen = tree.xpath('//strong/text()')
#print('ploegen: ', ploegen)