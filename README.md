# certification_generator
교육수료증을 생성해주는 자동화 프로그램입니다



생성하기를 클릭한 후 조건에 맞게 입력된 엑셀파일을 선택합니다
그럼 개개인의 수료증파일과 이를 한번에 인쇄할 수 있도록 통합된 수료증 파일이 생성됩니다

cert.docx라는 파일의 폼을 확인한 후 수정할 수 있는 부분은 수정이 가능합니다.(출력날짜 등등)
이를 활용하면 다른 폼에도 적용할 수 있겠죠?

처음 docx 모듈을 다운받으면 동아시아에 대한 폰트가 적용되지 않습니다.(궁서 등등)
이를 수정하기위해서 모듈을 살짝 수정해야합니다.

(참고 사이트)
https://github.com/python-openxml/python-docx/pull/576/files   한글폰트관련

https://stackoverflow.com/questions/24872527/combine-word-document-using-python-docx   여러개의 문서 하나로 통합하는 방법

https://stackoverflow.com/questions/35642322/pyinstaller-and-python-docx-module-do-not-work-together   pyinstaller가  docx를 인식하지못하는 에러 
