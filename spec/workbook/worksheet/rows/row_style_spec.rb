# frozen_string_literal: true

require 'spec_helper'

describe 'XlsxRow', '#stlye' do
  it 'row has no style' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/row_height_nil.xlsx')
    expect(xlsx.worksheets.first.rows[0].style).to be_nil
  end

  it 'row has some style' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/row_style/custom_row_style.xlsx')
    expect(xlsx.worksheets.first.rows[0].style).to be_a(OoxmlParser::Xf)
  end
end
