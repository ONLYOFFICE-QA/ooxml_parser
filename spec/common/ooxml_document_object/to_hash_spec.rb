require 'spec_helper'

describe 'to_hash' do
  it 'to_hash for simple object' do
    docx = OoxmlParser::Parser.parse('spec/common/ooxml_document_object/to_hash/text_in_paragraph.docx')
    expect(docx.elements.first.background_color.to_hash).to be_a(Hash)
  end

  it 'to_hash for docx paragraph' do
    docx = OoxmlParser::Parser.parse('spec/common/ooxml_document_object/to_hash/text_in_paragraph.docx')
    expect(docx.elements.first.to_hash).to be_a(Hash)
  end

  it 'to_hash for doc object' do
    docx = OoxmlParser::Parser.parse('spec/common/ooxml_document_object/to_hash/text_in_paragraph.docx')
    expect(docx.to_hash).to be_a(Hash)
  end

  it 'to_hash works for array' do
    docx = OoxmlParser::Parser.parse('spec/common/ooxml_document_object/to_hash/text_in_paragraph.docx')
    expect(docx.to_hash[:@elements_0]).to be_a(Hash)
  end
end
