require 'rspec'
require 'ooxml_parser'

describe 'File Path In Parsed sctructure' do
  it 'DocxStructure contains file path' do
    file_path = 'spec/docx_examples/elements/paragraph/character/text/tab_inside_run.docx'
    docx = OoxmlParser::DocxParser.parse_docx(file_path)
    expect(docx.file_path).to eq(file_path)
  end

  it 'XLSX Workbook contains file path' do
    file_path = 'spec/xlsx_examples/editor_specific_documents/ms_office_2013/simple_text.xlsx'
    xlsx = OoxmlParser::XlsxParser.parse_xlsx(file_path)
    expect(xlsx.file_path).to eq(file_path)
  end

  it 'PPTX presentation contains file path' do
    file_path = 'spec/pptx_examples/settings/measurement.pptx'
    pptx = OoxmlParser::PptxParser.parse_pptx(file_path)
    expect(pptx.file_path).to eq(file_path)
  end
end
