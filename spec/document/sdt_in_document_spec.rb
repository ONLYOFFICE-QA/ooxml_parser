require 'spec_helper'

describe 'Sdt in document' do
  it 'sdt_in_document.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/sdt/sdt_in_document.docx')
    expect(docx.elements.first.sdt_content.paragraphs.first.runs.first.text).to eq('Simple Test Text')
  end
end
