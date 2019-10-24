# frozen_string_literal: true

require 'spec_helper'

describe 'Table#table_columns' do
  it 'Table Columns data' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/tables/table_columns/table_columns_data.xlsx')
    expect(xlsx.worksheets.first.table_parts.first.table_columns.count).to eq(2)
    expect(xlsx.worksheets.first.table_parts.first.table_columns[0].id).to eq(1)
    expect(xlsx.worksheets.first.table_parts.first.table_columns[0].name).to eq('Column1')
    expect(xlsx.worksheets.first.table_parts.first.table_columns[0].totals_row_label).to eq('Summary')
    expect(xlsx.worksheets.first.table_parts.first.table_columns[1].totals_row_function).to eq('sum')
  end
end
