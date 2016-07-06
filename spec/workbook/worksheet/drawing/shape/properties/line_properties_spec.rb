require 'spec_helper'

describe 'My behaviour' do
  it 'Shape Stroke Size None' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/properties/line/stroke_width_none.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.shape.properties.line.stroke_size).to eq(0)
  end

  it 'Shape Stroke Size 6px' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/properties/line/stroke_width_6px.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.shape.properties.line.stroke_size).to eq(6)
  end

  it 'ShapeStrokeColor1' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/properties/line/stroke_color_1.xlsx')
    expect(xlsx.worksheets.first.drawings.first.shape.properties.line.color_scheme.color).to eq(OoxmlParser::Color.new(248, 203, 173))
  end

  it 'ShapeStrokeColor2' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/properties/line/stroke_color_2.xlsx')
    expect(xlsx.worksheets.first.drawings.first.shape.properties.line.color_scheme.color).to eq(OoxmlParser::Color.new(192, 0, 0))
  end
end
