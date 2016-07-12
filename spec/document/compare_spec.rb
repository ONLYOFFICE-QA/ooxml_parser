require 'spec_helper'

describe 'compare' do
  it 'Compare two similar docx - set 1' do
    docx1 = OoxmlParser::DocxParser.parse_docx('spec/document/compare/set_1/first.docx')
    docx2 = OoxmlParser::DocxParser.parse_docx('spec/document/compare/set_1/second.docx')
    expect(docx1).to eq(docx2)
  end

  it 'Compare two similar docx - set 2' do
    docx1 = OoxmlParser::DocxParser.parse_docx('spec/document/compare/set_2/first.docx')
    docx2 = OoxmlParser::DocxParser.parse_docx('spec/document/compare/set_2/second.docx')
    expect(docx1).to eq(docx2)
  end

  it 'compare_two_paragraphs' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/compare/compare_two_paragraphs.docx')
    expect(docx.elements[0]).to eq(docx.elements[1])
  end

  it 'compare_two_different_document' do
    docx1 = OoxmlParser::DocxParser.parse_docx('spec/document/compare/different_set/first.docx')
    docx2 = OoxmlParser::DocxParser.parse_docx('spec/document/compare/different_set/second.docx')
    expect(docx1).not_to eq(docx2)
  end
end
