require 'spec_helper'

describe 'Sparkline group' do
  it 'sparkline_group_type.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/extensions/sparkline_groups/sparkline_group/sparkline_group_type.xlsx')
    expect(xlsx.worksheets.first.extension_list[0].sparkline_groups[0].type).to eq(:column)
  end

  it 'sparkline_line_weight.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/extensions/sparkline_groups/sparkline_group/sparkline_line_weight.xlsx')
    expect(xlsx.worksheets.first.extension_list[0].sparkline_groups[0].line_weight).to eq(OoxmlParser::OoxmlSize.new(0.5, :point))
  end
end
