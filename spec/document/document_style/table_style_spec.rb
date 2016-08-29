require 'spec_helper'

describe 'TableStyle' do
  it 'table_style_height' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/table_style/table_style_height.docx')
    expect(docx.document_style_by_name('CustomTableStyle').table_row_properties.height.value).to eq(OoxmlParser::OoxmlSize.new(1440, :twip))
  end
end
