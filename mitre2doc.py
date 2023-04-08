import re
from bs4 import BeautifulSoup
import requests
from odf.opendocument import OpenDocumentText
from odf.style import Style, TextProperties
from odf.text import H, P, List, ListItem

def get_mitre(url):
    # define client and prepare header
    user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36'
    headers = {'User-Agent': user_agent}
    try: 
        # send get request to MITRE
        response = requests.get(url, headers=headers)
    except Exception as e:
        print('Error accessing webpage. Error Message: ', str(e))
    # print the response status code to the console
    print('Response: ', response)
    return response


def main():
    response = get_mitre('https://attack.mitre.org/techniques/enterprise/')
    try:
        # parse webpage
        website_content = BeautifulSoup(response.content, 'html.parser')
        # find table and its rows
        table_body = website_content.find('tbody')
        rows = table_body.find_all('tr')
        # create document basis and styles
        doc = OpenDocumentText()
        s = doc.styles
        # Heading
        title = Style(name="Title", family="paragraph")
        title.addElement(TextProperties(attributes={'fontsize':"28pt",'fontweight':"bold"}))
        # Headline H1
        h1 = Style(name="Heading 1", family="paragraph")
        h1.addElement(TextProperties(attributes={'fontsize':"20pt",'fontweight':"bold", 'fontfamily':"Liberation Sans"}))
        s.addElement(h1)
        # Headline H2
        h2 = Style(name="Heading 2", family="paragraph")
        h2.addElement(TextProperties(attributes={'fontsize':"14pt",'fontweight':"bold", 'fontfamily':"Liberation Sans"}))
        s.addElement(h2)
        # Text Body, used for the descriptions
        body = Style(name="Text Body Indent", family="paragraph")
        body.addElement(TextProperties(attributes={'fontsize': "12pt", 'fontfamily':" Liberation Sans"}))
        s.addElement(body)
        # List Style
        liststyle = Style(name="List Content", family="paragraph")
        liststyle.addElement(TextProperties(attributes={'fontsize':"12pt", 'fontfamily':"Liberation Sans"}))
        s.addElement(liststyle)
        # create headline and add it to document
        headline=H(outlinelevel=0, stylename=title, text="MITRE ATT&CK - Techniques and Subtechniques")
        doc.text.addElement(headline)
        # for each row displayed in the table on the webpage
        for entry in rows:
            # get the table data and clean it from html tags
            contents = entry.find_all('td')
            description = contents[len(contents)-1].text.replace('\n', '').replace('  ', '')
            name = contents[len(contents)-2].text.replace('\n', '')
            # set the heading = name of the MITRE Technique
            if entry.attrs['class'][0] == "technique":
                technique = name
                h_technique=P(stylename=h1, text=name + '\n')
                doc.text.addElement(h_technique)
                # for techniques also get the tactics from the technique page
                sub_page = entry.contents[1].contents[1].attrs['href']
                technique_page = get_mitre('https://attack.mitre.org/' + sub_page)
                technique_page = BeautifulSoup(technique_page.content, 'html.parser')
                # get tactics from webpage
                tactics = technique_page.find('div', {'id': 'card-tactics'})
                # clean extracted content
                tactics = tactics.text.replace('\n', '').replace('ⓘTactics:', '').replace('ⓘTactic:', '')
                tactics = re.sub(re.compile('[ ]{2,5}'), '', str(tactics))
            else:
                # Subtechnique with H2 Heading
                h_subTechnique=P(stylename=h2, text=technique + "-" + name )
                doc.text.addElement(h_subTechnique)
            # set the content of the paragraph = description
            p_desc = P(stylename=body, text=description)
            doc.text.addElement(p_desc)
            # items to be added
            list_items_to_add = ['Tactics: ' + tactics, 'Relevant? Yes/No']
            myList = List(stylename=liststyle)
            for item in list_items_to_add:
                # add list items
                li = ListItem()
                li.addElement(P(stylename = liststyle, text = item))
                myList.addElement(li)
                doc.text.addElement(myList)
        # save the document
        doc.save('MITRE ATT&CK.docx', addsuffix=False)
        print('Finished creating MITRE ATT&CK document file!')
    except Exception as e:
        print('Error creating document file. Error Message: ', str(e))
    

if __name__ == '__main__':
    main()