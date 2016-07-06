require 'spec_helper'

describe 'My behaviour' do
  it 'ShapeGrouping' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shapes_grouping/grouping_shape.pptx')
    expect(pptx.slides[0].elements[0]).to be_a_kind_of OoxmlParser::ShapesGrouping
  end

  it 'grouping_chart.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shapes_grouping/grouping_chart.pptx')
    elements = pptx.slides.first.elements
    expect(elements.first).to be_a_kind_of OoxmlParser::ShapesGrouping
    expect(elements[0].elements.size).to eq(2)
  end
end
