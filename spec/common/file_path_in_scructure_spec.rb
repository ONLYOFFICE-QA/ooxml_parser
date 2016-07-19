require 'spec_helper'

describe 'File Path In Parsed sctructure' do
  it 'DocxStructure contains file path' do
    file_path = 'spec/common/file_path/file_path.docx'
    docx = OoxmlParser::DocxParser.parse_docx(file_path)
    expect(docx.file_path).to eq(file_path)
  end

  it 'XLSX Workbook contains file path' do
    file_path = 'spec/common/file_path/file_path.xlsx'
    xlsx = OoxmlParser::XlsxParser.parse_xlsx(file_path)
    expect(xlsx.file_path).to eq(file_path)
  end

  it 'PPTX presentation contains file path' do
    file_path = 'spec/common/file_path/file_path.pptx'
    pptx = OoxmlParser::PptxParser.parse_pptx(file_path)
    expect(pptx.file_path).to eq(file_path)
  end

  it 'File path for common parser' do
    file_path = 'spec/common/file_path/file_path.docx'
    docx = OoxmlParser::Parser.parse(file_path)
    expect(docx.file_path).to eq(file_path)
  end
end
