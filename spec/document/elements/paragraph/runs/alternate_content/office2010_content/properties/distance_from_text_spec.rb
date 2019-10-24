# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'distance_from_text_in_emu' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/runs/alternate_content/office2010_content/properties/distance_from_text/distance_from_text_in_emu.docx')
    expect(docx.elements.first.character_style_array[2].alternate_content.office2010_content.properties.distance_from_text.left).to eq(OoxmlParser::OoxmlSize.new(914_400, :emu))
  end
end
