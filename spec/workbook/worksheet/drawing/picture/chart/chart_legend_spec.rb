require 'spec_helper'

describe 'My behaviour' do
  it 'Insert Chart Window | Check Legend Left Position without Overlay' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/legend/legend_left_position.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.graphic_frame.graphic_data.first.legend.position).to eq(:left)
    expect(drawing.graphic_frame.graphic_data.first.legend.overlay).to eq(false)
  end

  it 'Insert Chart Window | Check Legend Left Position with Overlay' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/legend/left_position_with_overlay.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.graphic_frame.graphic_data.first.legend.overlay).to eq(true)
  end

  describe 'Legend position' do
    it 'Chart - Legend - legend_left_overlay.xlsx' do
      xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/legend/legend_left_overlay.xlsx')
      drawing = xlsx.worksheets[0].drawings[0]
      expect(drawing.graphic_frame.graphic_data.first.legend.position_with_overlay).to eq(:left_overlay)
    end

    it 'Chart - Legend - legend_right_overlay.xlsx' do
      xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/legend/legend_right_overlay.xlsx')
      drawing = xlsx.worksheets[0].drawings[0]
      expect(drawing.graphic_frame.graphic_data.first.legend.position_with_overlay).to eq(:right_overlay)
    end

    it 'Chart - Legend - legend_right.xlsx' do
      xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/legend/legend_right.xlsx')
      drawing = xlsx.worksheets[0].drawings[0]
      expect(drawing.graphic_frame.graphic_data.first.legend.position_with_overlay).to eq(:right)
    end
  end
end
