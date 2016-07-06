require 'spec_helper'

describe 'My behaviour' do
  describe 'libreoffice' do
    it 'simple_text' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/editor_specific_documents/libreoffice/simple_text.docx')
      expect(docx.elements.first.character_style_array.first.text).to eq('This is a test')
    end
  end

  describe 'ms_office_2013' do
    it 'simple_text' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/editor_specific_documents/ms_office_2013/simple_text.docx')
      expect(docx.elements.first.character_style_array.first.text).to eq('This is a test')
    end
  end
end
