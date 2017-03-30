require 'spec_helper'

describe 'My behaviour' do
  it 'Chart Without Title' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/title/chart_without_title.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.graphic_frame.graphic_data.first.title.elements.first).to be_nil
  end

  it 'Chart With Title' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/title/chart_with_titile.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.graphic_frame.graphic_data.first.title.elements.first.characters.first.text).to eq('Custom Title')
  end

  it 'Chart With Empty Title' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/title/chart_with_empty_title.xlsx')
    drawing = xlsx.worksheets[0].drawings[0]
    expect(drawing.graphic_frame.graphic_data.first.title.elements.first.characters.first).to be_nil
  end

  it 'ChartTitle' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/picture/chart/title/chart_custom_title.xlsx')
    expect(xlsx.worksheets[0].drawings[0].graphic_frame.graphic_data.first.title.elements.first.characters[0].text).to eq('CustomChartTitle')
  end
end
