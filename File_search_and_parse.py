
import docx2txt
import os
import glob



feath_dir=r"D:\Documents\Nick test folder\Nick Journal and Thoughts_v2"
search_word="t"
search_list=[]
file_list=[]
a=r"D:\Documents\Nick test folder\Nick Journal and Thoughts_v2\Goals20190601_test.docx"
x = int()
y = int()
file_limit=2
paragraph_limit=5
para_sentence_limit=4
file_counter=int()

for a in glob.glob(feath_dir+"/"+"*.docx"):
    para_counter=0
    file_counter+=1
    if a.count(search_word)>0:
        paragraph_list=docx2txt.process(a)
        # b is each paragraph in a file
        for b in paragraph_list.split('\n'):
            b_sentence_count = 0
            para_counter+=1
            print(file_counter, para_counter, b_sentence_count)
            if b.count(search_word)>0:
                if b.count(".")<=1:
                    file_list.append(os.path.basename(a))
                    search_list.append(b)
                    x=len(file_list)
                    y=len(search_list)
                # establishes the paragraph has more than one sentence with elif below
                elif b.count(".")>1:
                    while b.count(".")>0:
                        # print("elif b.count(.)>1 is " + str(b.count(".")))
                        b_subsentence = b[0:b.find(".") +1]
                        b_sentence_count +=1
                        # establishes the search_word is not in b_subsentence
                        if b_subsentence.count(search_word) == 0:
                            # make b everything b was except b_subsentce
                            b = b[len(b_subsentence):]
                        # search_word is in b_subsentence
                        elif b_subsentence.count(search_word)>0:
                            search_list.append(b_subsentence)
                            file_list.append(os.path.basename(a))
                            x = len(file_list)
                            y = len(search_list)
                            b = b[len(b_subsentence):]
                            # print(file_list[x - 1], x, search_list[y - 1], para_counter)
                            # ***********LIMIT number of SENTENCES in paragraph***********
                            if b_sentence_count >= para_sentence_limit:
                                break
            # ***********LIMIT number of PARAGRAPHS in file***********
            elif para_counter>=paragraph_limit:
                print('this filed_ 2')
                break
        # ***********LIMIT number of Files***********
        if file_counter>=file_limit:
            break