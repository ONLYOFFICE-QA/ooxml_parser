require 'spec_helper'

describe 'My behaviour' do
  it 'document_background' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/background/document_background.docx')
    expect(docx.background.color1).to eq(OoxmlParser::Color.new(255, 255, 255))
  end

  it 'document_background_fill' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/background/document_background_fill.docx')
    expect(docx.background.file_reference.content.length).to be > 1000
  end
end
