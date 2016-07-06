require 'spec_helper'

describe 'document properties' do
  it 'Page Count' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_properties/page_count.docx')
    expect(docx.document_properties.pages).to eq(2)
  end

  it 'Word Count' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_properties/word_count.docx')
    expect(docx.document_properties.words).to eq(1)
  end

  it 'no_app_xml_file' do
    expect do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_properties/no_app_xml_file.docx')
      expect(docx.document_properties.pages).to be_nil
    end.to output(%r{no 'docProps\/app.xml'}).to_stderr
  end

  it 'no_word_statistic' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_properties/no_word_statistic.docx')
    expect(docx.document_properties.pages).to be_nil
    expect(docx.document_properties.words).to be_nil
  end
end
