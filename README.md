# certificate-generator
- generate certificates by passing a json file

## Requirements:
- Python 3+

## Support:
- generate Microsoft Word

## Setup:
- 

## How to use:
1. Initially run the program to setup
```shell
python main.py || cd bin && generate-certificate
```
2. See the sample.json file
3. Edit the data.json file 

## Flow: 
1. prompt, if first run then with setup, or if with params --setup
2. choose either setup or generate:
   1. setup
      1. add field
      2. edit field
         1. loop then choose number to edit
      3. remove field
         1. loop choose number to remove
      4. clear fields 
   2. generate
      1. read the data.json file 
      2. get the template/template-word.docx
      3. based on data there, loop through the data also based on accepted fields
         1. field on the fields
         2. save 

## Preview:
- https://www.loom.com/share/7dba772b26dd4152bfe7c989f1294af3?sid=44968d1c-50db-476b-91e7-fcf2231a7650
