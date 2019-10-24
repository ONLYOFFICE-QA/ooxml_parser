# frozen_string_literal: true

require 'spec_helper'

describe 'Connection Shape' do
  it 'Connection Shape exists' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/connection_shape/connection_shape_exists.xlsx')
    expect(xlsx.worksheets.first.drawings.first.shape).to be_a(OoxmlParser::ConnectionShape)
  end
end
