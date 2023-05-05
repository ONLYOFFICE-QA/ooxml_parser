# frozen_string_literal: true

require 'spec_helper'

describe 'Sdt in document' do
  sdt_with_table = OoxmlParser::Parser.parse('spec/document/sdt/sdt_with_table.docx')

  it 'sdt_in_document.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/sdt/sdt_in_document.docx')
    expect(docx.elements.first.sdt_content.paragraphs.first.runs.first.text).to eq('Simple Test Text')
  end

  it 'sdt_with_table.docx' do
    expect(sdt_with_table.elements.first.sdt_content.elements.first).to be_a(OoxmlParser::Table)
  end

  it 'SDTContent#table' do
    expect(sdt_with_table.elements.first.sdt_content.tables.first).to be_a(OoxmlParser::Table)
  end

  it 'sdt_element_parent.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/sdt/std_element_parent.docx')
    expect(docx.elements.first.sdt_content.parent.parent).to be_a(OoxmlParser::DocumentStructure)
  end
end
