require 'spec_helper'

describe 'My behaviour' do
  it 'Shape Naming' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/properties/name/shape_naming.pptx')
    expect(pptx.slides[0].elements[-1].shape_properties.preset.name).to eq(:rect)
  end

  it 'shape_group_naming.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/properties/name/shape_group_naming.pptx')
    expect(pptx.slides.first.elements.first.elements.first.shape_properties.preset.name).to eq(:ellipse)
    expect(pptx.slides.first.elements.first.elements.last.shape_properties.preset.name).to eq(:mathPlus)
  end
end
