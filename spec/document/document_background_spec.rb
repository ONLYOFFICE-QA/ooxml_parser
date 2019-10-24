# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'document_background' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/background/document_background.docx')
    expect(docx.background.color1).to eq(OoxmlParser::Color.new(255, 255, 255))
  end

  it 'document_background_fill' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/background/document_background_fill.docx')
    expect(docx.background.fill.file.content.length).to be > 1000
  end

  it 'document_background_fill_no_type' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/background/document_background_fill_no_type.docx')
    expect(docx.background.fill.color2).to eq(OoxmlParser::Color.new(0, 0, 0))
  end
end
