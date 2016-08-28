require 'spec_helper'

describe 'My behaviour' do
  it 'chart_transform_size' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/transform/chart_transform_size.pptx')
    transform = pptx.slides.last.elements.last.transform
    expect(transform.extents.x).to eq(OoxmlParser::OoxmlSize.new(8.4, :centimeter))
    expect(transform.extents.y).to eq(OoxmlParser::OoxmlSize.new(5, :centimeter))
  end
end
