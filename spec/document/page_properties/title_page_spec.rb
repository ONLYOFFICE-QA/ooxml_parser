require 'spec_helper'

describe 'Title Page' do
  it 'title_page_default.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/title_page/title_page_default.docx')
    expect(docx.elements.first.paragraph_properties.section_properties.title_page).to be_falsey
  end

  it 'title_page_enabled.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/title_page/title_page_enabled.docx')
    expect(docx.elements.first.paragraph_properties.section_properties.title_page).to be_truthy
  end
end
