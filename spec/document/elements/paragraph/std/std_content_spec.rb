require 'spec_helper'

describe 'StdContent' do
  it 'check std_content_text' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/std/std_content/std_content_text.docx')
    expect(docx.elements.first.sdt.sdt_content.runs.first.text).to eq('Simple Test Text')
  end

  it 'sdt_properties_tag.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/std/std_content/sdt_properties_tag.docx')
    expect(docx.elements.first.sdt.properties.tag.value).to eq('ContentControlTag')
  end

  it 'sdt_properties_lock.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/std/std_content/sdt_properties_lock.docx')
    expect(docx.elements.first.sdt.properties.lock.value).to eq(:sdt_content_locked)
  end
end
