require 'spec_helper'

describe OoxmlParser::CommonNonVisualProperties do
  it 'cnvsp_title' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/non_visual_properties/common_non_visual_properties/cnvsp_title.xlsx')
    expect(xlsx.worksheets.first.drawings.first.shape.non_visual_properties.common_properties.title).to eq('Simple Test Text')
  end

  it 'cnvsp_description' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/non_visual_properties/common_non_visual_properties/cnvsp_description.xlsx')
    expect(xlsx.worksheets.first.drawings.first.shape.non_visual_properties.common_properties.description).to eq('SimpleTestText')
  end
end
