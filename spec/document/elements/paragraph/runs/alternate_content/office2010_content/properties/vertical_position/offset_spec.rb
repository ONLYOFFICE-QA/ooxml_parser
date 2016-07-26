require 'spec_helper'

describe 'My behaviour' do
  it 'should do something' do
    OoxmlParser.configure do |config|
      config.units = :emu
    end

    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/runs/alternate_content/office2010_content/properties/vertical_position/offset/offset_in_other_unit.docx')
    expect(docx.elements.first.character_style_array[2].alternate_content.office2010_content.properties.vertical_position.offset).to eq(914_400)
  end

  after do
    OoxmlParser.reset
  end
end
