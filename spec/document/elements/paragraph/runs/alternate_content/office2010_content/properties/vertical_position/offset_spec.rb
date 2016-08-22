require 'spec_helper'

describe 'My behaviour' do
  it 'offset_in_other_unit' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/runs/alternate_content/office2010_content/properties/vertical_position/offset/offset_in_other_unit.docx')
    expect(docx.elements.first.character_style_array[2].alternate_content.office2010_content.properties.vertical_position.offset).to eq(OoxmlParser::OoxmlSize.new(914_400, :emu))
  end

  after do
    OoxmlParser.reset
  end
end
