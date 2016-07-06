require 'spec_helper'

describe 'My behaviour' do
  it 'chart_stroke_no_line.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/graphic_data/chart/shape_properties/line/chart_stroke_no_line.pptx')
    drawing = pptx.slides[0].elements.last
    expect(drawing.graphic_data.first.shape_properties.line.width).to eq(0)
  end

  it 'chart_stroke_0.5.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/graphic_data/chart/shape_properties/line/chart_stroke_0.5.pptx')
    drawing = pptx.slides[0].elements.last
    expect(drawing.graphic_data.first.shape_properties.line.width).to eq(0.5)
  end

  it 'chart_borders.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/graphic_data/chart/shape_properties/line/chart_borders.pptx')
    expect(pptx.slides[0].elements.first.graphic_data.first.shape_properties.line.width).to eq(0)
    expect(pptx.slides[1].elements.first.graphic_data.first.shape_properties.line.color).to be_nil
    expect(pptx.slides[2].elements.first.graphic_data.first.shape_properties.line.width).to eq(0.5)
    expect(pptx.slides[2].elements.first.graphic_data.first.shape_properties.line.color).not_to be_nil
  end
end
