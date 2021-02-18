# frozen_string_literal: true

require 'spec_helper'

describe 'Runs Font Size' do
  it 'font_size_float' do
    docx = OoxmlParser::DocxParser.parse_docx('/home/lobashov/temp/set_shd.js20210218-30830-7q608r.docx')
    expect(docx.elements.first.background_color).to eq(OoxmlParser::Color.new(0, 255, 0))
  end
end
