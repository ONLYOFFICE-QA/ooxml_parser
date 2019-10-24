# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 't_formula' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/formula/t_formula.xlsx')
    expect(xlsx.worksheets[0].rows[0].cells[0].formula.value).to eq('T(1)')
  end
end
