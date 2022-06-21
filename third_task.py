import os
from bs4 import BeautifulSoup

current_directory = os.path.abspath(os.path.dirname('Data_Repositories_Test.docx'))


def wrap(to_wrap, wrap_in):
    content = to_wrap.replace_with(wrap_in)
    wrap_in.append(content)


path = os.path.join(current_directory, 'Explainers')
files = os.listdir(path)
for one_file in files:
    file = open(os.path.join(path, one_file), 'r', encoding='utf-8')
    contents = file.read()
    soup = BeautifulSoup(contents, 'xml')

    for paragraph in soup.find_all('p'):
        paragraph.name = 's'

    for sparagraph in soup.find_all('s'):
        wrap(sparagraph, soup.new_tag("p"))

    for tabled in soup.find_all('td'):
        tabled.name = 'ss'

    for sstabled in soup.find_all('ss'):
        wrap(sstabled, soup.new_tag("td"))
        sstabled.name = 's'

    with open(os.path.join(current_directory, 'Output', 'third_task', one_file), 'w', encoding='utf-8') as file:
        file.write(str(soup))
