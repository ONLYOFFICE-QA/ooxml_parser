# frozen_string_literal: true

require 'spec_helper'

describe 'Empty borders properties' do
  it 'empty_borders_properties.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/paragraph_properties/paragraph_borders/empty_borders_properties.docx')
    expect(docx.elements.first.borders.border_visual_type).to eq(:none)
  end
end
