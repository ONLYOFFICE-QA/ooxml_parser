require 'spec_helper'

describe 'Universal Method to parse file' do
  it 'DocxStructure is parsed' do
    file_path = 'spec/common/file_type_recognize/simple_docx.docx'
    docx = OoxmlParser::Parser.parse(file_path)
    expect(docx).to be_a(OoxmlParser::DocumentStructure)
  end

  it 'XLSX Workbook is parsed' do
    file_path = 'spec/common/file_type_recognize/simple_xlsx.xlsx'
    xlsx = OoxmlParser::Parser.parse(file_path)
    expect(xlsx).to be_a(OoxmlParser::XLSXWorkbook)
  end

  it 'PPTX presentation is parsed' do
    file_path = 'spec/common/file_type_recognize/simple_pptx.pptx'
    pptx = OoxmlParser::Parser.parse(file_path)
    expect(pptx).to be_a(OoxmlParser::Presentation)
  end

  it 'Unrecognized Zip Archive' do
    expect do
      file_path = 'spec/common/file_type_recognize/simple_zip.zip'
      zip = OoxmlParser::Parser.parse(file_path)
      expect(zip).to be_nil
    end.to output(/simple zip file without OOXML content/).to_stderr
  end
end
