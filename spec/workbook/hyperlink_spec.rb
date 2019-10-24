# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'hyperlink_internal.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/hyperlinks/hyperlink_internal.xlsx')
    expect(xlsx.worksheets[0].hyperlinks[0].link).to eq(OoxmlParser::Coordinates.new(10, 'F'))
    expect(xlsx.worksheets[0].rows[0].cells.first.text).to eq('yandex')
    expect(xlsx.worksheets[0].hyperlinks[0].tooltip).to eq('go to www.yandex.ru')
  end

  it 'hyperlink_empty_id.xlsx' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/hyperlinks/hyperlink_empty_id.xlsx')
    expect(xlsx.worksheets[1].drawings.first.shape.non_visual_properties.common_properties.on_click_hyperlink.id).to be_empty
  end
end
