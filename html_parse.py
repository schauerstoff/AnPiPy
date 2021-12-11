from bs4 import BeautifulSoup
from PyInquirer import prompt, Separator

# Returns each li from ol{ul{li}}, PyInquirer.separator between each ol-element, as a list
def flatten_and_seperate(list):
        # flatten 2D-list and add seperator between elements
        return ', '.join([j for sub in list for j in sub])

def parse_html_list(html_code):
    res = []
    olList = html_code.find_all('ul')
    if(olList):
        for elem in olList:
            soup = BeautifulSoup(str(elem), features="lxml")
            ulList = soup.find_all('li')
            for word in ulList:
                word = str(word).replace('<li>', '').replace(
                    '''</li>''', '').replace('\'', '').strip()
                res.append(word)
            res.append(Separator('___________________'))
        return res
    # no nested ul-List, check only one ol
    else:
        oneList = html_code.find_all('li')
        if(oneList):
            for word in oneList:
                word = str(word).replace('<li>', '').replace(
                    '</li>', '').replace('\'', '').strip()
                res.append(word)
            return res
        # only one word!
        else:
            oneWord = html_code.find_all('p')
            for word in oneWord:

                word = str(word).replace('<p>', '').replace(
                    '</p>', '').replace('\'', '').strip()
                res.append(word)
            return res


# html = BeautifulSoup("<ol><li> <ul><li>light'ning</li><li>thunder</li><li>thunderbolt</li></ul></li><li> <ul><li>god of thunder</li><li>god of lightning</li></ul></li><li> <ul><li>anger</li><li>fit of anger</li></ul></li><li> <ul><li>lightning</li><li>thunder</li><li>thunderbolt</li><li>god of thunder</li><li>god of lightning</li><li>anger</li><li>fit of anger</li></ul></li></ol>", features="lxml")
# words = parse_html_list(html)
# print(flatten_and_seperate(words))

# html = BeautifulSoup("tuna", features="lxml")
# words = parse_html_list(html)
# print(words)

# html = BeautifulSoup(" to suit one&#x27;s taste", features="lxml")
# #html = BeautifulSoup("to suit ones taste", features="lxml")
# words = parse_html_list(html)
# print(words)

# html = BeautifulSoup(
#     "<ol><li> vendin'g machine</li><li> vending machine</li></ol>", features="lxml")
# words = parse_html_list(html)
# print(words)

word = ['tuna']
#print(flatten_and_seperate(word))
print(isinstance(word[0], list))

word = [['lightning', 'thunder']]
#print(flatten_and_seperate(word))
print(isinstance(word[0], list))