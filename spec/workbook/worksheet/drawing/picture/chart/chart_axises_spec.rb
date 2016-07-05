require 'spec_helper'

describe 'My behaviour' do
  it 'Insert Chart Window | Check Set Display X Axis' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/axises/axis_display_x.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.picture.chart.axises.first.title).not_to be_nil
  end

  it 'Insert Chart Window | Check Set Not Display X Axis' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/axises/axis_not_display_x.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.picture.chart.axises.first.title).to be_nil
  end

  describe 'Axises' do
    it 'Chart - Axises - axis_display_none.xlsx' do
      xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/axises/axis_display_none.xlsx')
      drawing = xlsx.worksheets[0].drawings[0]
      expect(drawing.picture.chart.axises[1].display).to eq(false)
    end

    it 'Chart - Axises - axis_display_overlay.xlsx' do
      xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/axises/axis_display_overlay.xlsx')
      drawing = xlsx.worksheets[0].drawings[0]
      expect(drawing.picture.chart.axises[1].display).to eq(true)
    end

    it 'Chart - Axises - axis_position.xlsx' do
      xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/axises/axis_position.xlsx')
      drawing = xlsx.worksheets[0].drawings[0]
      expect(drawing.picture.chart.axises[1].position).to eq(:left)
    end
  end
end
