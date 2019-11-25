# frozen_string_literal: true

require 'spec_helper'

describe 'Workbook#all_formula_values' do
  it 'value_of_empty_f.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/all_formula_values_spec/value_of_empty_f.xlsx')
    expect(xlsx.all_formula_values).to eq(['6.0', 'Error in calculation', '0.0', '6.0', '6.0', 'Error in calculation', '0.0'])
  end

  it 'all_formulas_values' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/all_formula_values_spec/all_formulas_values.xlsx')
    expect(xlsx.all_formula_values.length).to eq(12)
  end

  it 'all_formulas_values should not return values of empty formulas' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/all_formula_values_spec/empty_formulas.xlsx')
    expect(xlsx.all_formula_values).to eq(['1.0'])
  end

  it 'all_formulas_values correct for complex numbers' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/all_formula_values_spec/complex_formula_valuse.xlsx')
    expect(xlsx.all_formula_values).to eq(['3.0+4.0i', '3.0+4.0i', '0.0+1.0i', '1.0'])
  end
end
