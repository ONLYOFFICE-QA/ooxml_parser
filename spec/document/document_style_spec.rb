# frozen_string_literal: true

require 'spec_helper'

describe 'document style' do
  it 'Do not crash if there is no styles.xml' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/no_styles_xml.docx')
    expect(docx).to be_with_data
  end

  it 'New Paragraph Document Style' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/new_paragraph_style.docx')
    expect(docx.document_styles.last.name).to eq('NewParagraphStyle')
  end

  it 'Paragraph Document Visible Style' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/style_visibility.docx')
    expect(docx.document_style_by_name('Heading 8')).to be_visible
    expect(docx.document_style_by_name('Footer')).not_to be_visible
  end

  it 'New Paragraph Document Style: method "style_exist?"' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/style_exists.docx')
    expect(docx).to be_style_exist('NewParagraphStyle')
  end

  it 'New Paragraph Document Style: method "style_exist?", negative' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/style_exists.docx')
    expect(docx).not_to be_style_exist('ThisStyleIsNotExist')
  end

  it 'paragraph_style_id_word' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/paragraph_style_id_word.docx')
    expect(docx.document_styles.first.style_id).to eq('Normal')
  end

  it 'table_document_style' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/table_document_style.docx')
    expect(docx.document_style_by_name('Table Grid').table_properties).to be_a(OoxmlParser::TableProperties)
  end

  it 'table_style_properties_document_style' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/table_style_properties_document_style.docx')
    expect(docx.document_style_by_name('Lined').table_style_properties_list.first).to be_a(OoxmlParser::TableStyleProperties)
  end

  it 'table_style_properties_type' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/table_style_properties_type.docx')
    expect(docx.document_style_by_name('Lined').table_style_properties_list.first.type).to eq(:band1Horz)
  end

  it 'document_style_based_on_style' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/document_style_based_on_style.docx')
    expect(docx.document_style_by_name('CustomTableStyle').based_on_style.name).to eq('Bordered - Accent 5')
  end

  it 'document_style_default' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/document_style_default.docx')
    expect(docx.document_style_by_name('Default Paragraph Font').default).to be_truthy
    expect(docx.document_style_by_name('Title').default).to be_falsey
  end
end
