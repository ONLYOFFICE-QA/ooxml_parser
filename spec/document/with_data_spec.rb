# frozen_string_literal: true

require 'spec_helper'

describe 'with_data' do
  it 'no_data' do
    docx = OoxmlParser::Parser.parse('spec/document/with_data/no_data.docx')
    expect(docx).not_to be_with_data
  end

  it 'several_empty_paragraphs' do
    docx = OoxmlParser::Parser.parse('spec/document/with_data/several_empty_paragraphs.docx')
    expect(docx).not_to be_with_data
  end

  it 'text_in_paragraph' do
    docx = OoxmlParser::Parser.parse('spec/document/with_data/text_in_paragraph.docx')
    expect(docx).to be_with_data
  end

  it 'shape' do
    docx = OoxmlParser::Parser.parse('spec/document/with_data/shape.docx')
    expect(docx).to be_with_data
  end

  it 'header' do
    docx = OoxmlParser::Parser.parse('spec/document/with_data/header.docx')
    expect(docx).to be_with_data
  end

  it 'page_break' do
    docx = OoxmlParser::Parser.parse('spec/document/with_data/page_break.docx')
    expect(docx).to be_with_data
  end

  it 'section_break' do
    docx = OoxmlParser::Parser.parse('spec/document/with_data/section_break.docx')
    expect(docx).to be_with_data
  end

  it 'paragraph_without_paragraph_properties' do
    docx = OoxmlParser::Parser.parse('spec/document/with_data/paragraph_without_paragraph_properties.docx')
    expect(docx).to be_with_data
  end
end
