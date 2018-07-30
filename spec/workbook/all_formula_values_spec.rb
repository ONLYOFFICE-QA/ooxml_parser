require 'spec_helper'

describe 'Workbook#all_formula_values' do
  it 'value_of_empty_f.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/all_formula_values_spec/value_of_empty_f.xlsx')
    expect(xlsx.all_formula_values).to eq(['6.0', 'Error in calculation', '0.0', '6.0', '6.0', 'Error in calculation', '0.0'])
  end
end
