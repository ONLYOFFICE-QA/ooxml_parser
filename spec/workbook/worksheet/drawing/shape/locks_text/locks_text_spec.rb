# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'Shape has locks_text' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/locks_text/locks_text.xlsx')
    expect(xlsx.worksheets.first.drawings.first.shape.locks_text).to be_falsey
  end
end
