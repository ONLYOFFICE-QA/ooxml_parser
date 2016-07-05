require 'spec_helper'

describe 'document' do
  describe 'background' do
    it 'document_background' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/background/document_background.docx')
      expect(docx.background.color1).to eq(OoxmlParser::Color.new(255, 255, 255))
    end

    it 'document_background_fill' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/background/document_background_fill.docx')
      expect(File.exist?(docx.background.image)).to be_truthy
    end
  end

  describe 'editor_specific_documents' do
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

  it 'Check Error for Empty zip docx' do
    expect { OoxmlParser::DocxParser.parse_docx('spec/document/empty_zip_docx.docx') }.to raise_error(LoadError)
  end
end
