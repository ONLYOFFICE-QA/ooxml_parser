require 'spec_helper'

describe 'StdContent' do
  it 'check std_content_text' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/std/std_content/std_content_text.docx')
    expect(docx.elements.first.sdt.sdt_content.runs.first.text).to eq('Simple Test Text')
  end
end
