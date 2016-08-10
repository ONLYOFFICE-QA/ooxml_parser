require 'spec_helper'

describe 'document style' do
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
    expect(docx.style_exist?('NewParagraphStyle')).to be_truthy
  end

  it 'New Paragraph Document Style: method "style_exist?", negative' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/style_exists.docx')
    expect(docx.style_exist?('ThisStyleIsNotExist')).to be_falsey
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
end
