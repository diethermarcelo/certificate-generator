import docx
import json

def create_doc():
    template_path = './template'
    template_filename = 'template.docx'
    template_fullpath = f"{template_path}/{template_filename}"
    doc = docx.Document(template_fullpath)
    return doc

def replace_string(data):
    doc = create_doc()
    for p in doc.paragraphs:
        terms = data.keys()
        for term in terms:
            if term in p.text:
                replace_term = "{{" + term + "}}"
                p.text = p.text.replace(replace_term, data[term])
    return doc
def save_doc(doc, data, increment):
    current_file_name = data['file_name'] or "dest-{inc}".format(inc=increment)
    file_name = "./dist/{current_file_name}.docx".format(current_file_name=current_file_name)
    print("Creating {file_name}...".format(file_name=file_name))
    doc.save(file_name)

def main():
    try:
        f = open('./sample.json')
        full_data = json.load(f)
        i = 0
        for data in full_data:
            doc = replace_string(data)
            save_doc(doc, data, i)
            i += 1
    except KeyError:
        print('please input a file_name attribute on your sample.json file')
    except:
        print('problem occured. please try again later')
    return 1

if __name__ == '__main__':
    main()




