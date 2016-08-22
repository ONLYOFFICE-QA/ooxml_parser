require 'spec_helper'

describe 'My behaviour' do
  it 'emu_units_of_measurement' do
    OoxmlParser.configure do |config|
      config.units = :emu
    end
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/properties/indents/emu_units_of_measurement.docx')
    expect(docx.elements.first.paragraph_properties.indent.right_indent).to eq(2880)
  end

  after do
    OoxmlParser.reset
  end
end
