import sys, os
import zipfile
from tkinter import *
from tkinter import filedialog
import xml.etree.ElementTree as ET

# 파일읽어서 docx -> 폴더 zip 압축해제
def unzip(file_path) :
    try:
        print('### unzipping : ' + file_path)
        zipdocx = zipfile.ZipFile(file_path) # zipfile 객체생성
        
        target_path = os.path.join(os.getcwd(), 'zip', file_path.split('/')[-1]) # zip에 압축해제 (./zip/파일명)
        zipdocx.extractall(target_path) 
    except :
        print('zip압축해제 Error')
    else :
        print('zip압축해제 Success')
        return target_path # zip path
    finally :
        zipdocx.close()

# 해제한 폴더중 ./word/document.xml 파일선택
def get_documentxml(zip_path) :
    print('### Select document.xml from ' + zip_path)
    xml_path = os.path.join(zip_path, 'word', 'document.xml') # documnet.xml path
    return xml_path


# 선택한 파일중 xml <w:p> <w:r> <w:t>.text 파싱
def parse_content(xml_path) :
    print('### read xml file : ' + xml_path)
    tree = ET.parse(xml_path) # xml read
    root = tree.getroot() # root == w:document
    
    # namespace 설정
    ns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    # namespace 설정
    namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    prefix_namespace = '{' + namespace + '}'

    print('------------------------\n')
    

    for p in root.findall('./w:body/w:p', ns) : # para
        print(p)
        for r in p.findall('./w:r', ns) : # r
            print('\t', end ='')
            print(r)
            for t in r.findall('./w:t', ns) : # t
                print('\t\t', end ='')
                print(t)

    # 이름변환
    # 나이변환
    print('------------------------\n') 

# 선택한 파일중 xml 모든 <w:p> 파싱
def parse_p(xml_path):
    # namespace 설정
    # ns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    # namespace 설정
    namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    prefix_namespace = '{' + namespace + '}'
    ET.register_namespace('', namespace) # 파싱전에 Elementtree에 namespace 등록

    print('### read xml file : ' + xml_path)
    tree = ET.parse(xml_path) # xml read
    root = tree.getroot() # root == w:document

    print('p------------------------START\n')
    
    p_elem = []
    for p in root.iter(prefix_namespace+"p") : 
        p_elem.append(p)
        
    print(p_elem)
    
    print('p------------------------END\n')
    
    return p_elem # 모든 p_elem

# p_elem에서 t_elem 가져오기
def parse_t(xml_path, p_elem):
    # namespace 설정
    # ns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    # namespace 설정
    namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    prefix_namespace = '{' + namespace + '}'
    ET.register_namespace('', namespace) # 파싱전에 Elementtree에 namespace 등록

    print('### read xml file : ' + xml_path)
    tree = ET.parse(xml_path) # xml read
    root = tree.getroot() # root == w:document

    t_elem = {} # {p : t_elems}
    for p in p_elem : 
        t_elem[p] = []
        for t in p.iter(prefix_namespace+"t") :
            t_elem[p].append(t)

    # 출력    
    print('T------------------------START\n')
    print_elems(t_elem)
    print('T------------------------END\n')

    return t_elem

# 위치가져오기

# t_elem 에서 {variable} 찾기 # {variable}{variable} 붙이면안됨 > 무조건 {} 사이에는 space
def get_variable(t_elem) :
    variable_p = {} # {}포함된 p 추출
    for ptag, ttag in t_elem.items() :
        for ts in ttag :
            if ts.text.startswith('{') :  # '{'로 시작하는 원소가 잇을경우
                variable_p[ptag] = ttag
                break
    # 출력
    print('get variable------------------------START\n')
    print_elems(variable_p)
    print('get variable------------------------END\n')

    return variable_p

# {variable}가 있는 <p>에서 설정한 {}가 있는지 확인 -> 위치extract
def extract_variable(variable_p) :
    variable = ('{name}','{univ}', '{age}') # 파싱할 variable
    extract_variable = {} # 변경할 variable position이 담긴 dict
    
    # 출력
    print('extract variable------------------------START\n')

    for ptag, ttag in variable_p.items() : # line 가져오기
        temp = []
        for t in ttag : # line별 elems
            if t.text.startswith('{') :
                temp = []
            
            temp.append(t) # temp.append(t.text)
            
            if temp[-1].text.endswith('}') : # 마지막 원소가 }로 끝나면
                # text가져오기
                
                temp_txt = list(map(lambda elem : elem.text, temp)) # temp[] : element객체 > text 변환
                
                if ''.join(temp_txt) in variable : # {variable} 체크
                    print(''.join(temp_txt)) # 문서에서 보이는 txt
                    print(temp) # {} elem객체
                    print(temp_txt) # {} elem text
                    # dict append
                
    
    print('extract variable------------------------START\n')

# elem dict 출력

def print_elems(elem_dict) : # {ptag : [ttag]}
    for ptag, ttag in elem_dict.items() :
        print(ptag, end=' : ')
        for ts in ttag :
            print(ts.text, end=' / ')
        print('\n\n')


# 변경된 파일 폴더 -> docx zip 압축
def zip(complete_path):
    f = zipfile.ZipFile('archive.zip','w',zipfile.ZIP_DEFLATED)
    startdir = complete_path
    for dirpath, dirnames, filenames in os.walk(startdir):
        for filename in filenames:
            f.write(os.path.join(dirpath,filename))
    f.close()
    #docx_file = zipfile.ZipFile('./update/{}'.format(complete_path.split('/')[-1]), 'w')
    

if __name__ == "__main__":
#    file_path = filedialog.askopenfilename(initialdir='./')
#    zip_path = unzip(file_path) # 압축해제
#    xml_path = get_documentxml(zip_path) # 해제 후 document.xml select
    p_elem = parse_p("./zip/naver_resume.docx/word/document.xml") # p 파싱
    t_elem = parse_t("./zip/naver_resume.docx/word/document.xml", p_elem) # t 파싱
    variable_p = get_variable(t_elem) # t elems 에서 { } 파싱
    extract_variable(variable_p)
#    zip('C:/Users/HyeonGyu/Desktop/working/zip/이름.docx')
    
    

