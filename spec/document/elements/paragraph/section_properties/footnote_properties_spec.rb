# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::FootnoteProperties do
  it 'footnote_properties position' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/section_properties/footnote_properties/footnote_position.docx')
    expect(docx.page_properties.footnote_properties.position.value).to eq(:beneath_text)
  end

  it 'footnote_properties numbering format' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/section_properties/footnote_properties/footnote_position.docx')
    expect(docx.page_properties.footnote_properties.numbering_format.value).to eq(:decimal)
  end

  it 'footnote_properties numbering restart' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/section_properties/footnote_properties/footnote_position.docx')
    expect(docx.page_properties.footnote_properties.numbering_restart.value).to eq(:continuous)
  end

  it 'footnote_properties numbering start' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/section_properties/footnote_properties/footnote_position.docx')
    expect(docx.page_properties.footnote_properties.numbering_start.value).to eq(1)
  end
end
