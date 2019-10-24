# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'ShapesRuns' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/worksheet/drawing/shape/text_body/paragraphs/properties/spacing/spacing_after_ooxmlsize.xlsx')
    expect(xlsx.worksheets.first.drawings.first.shape.text_body.paragraphs.first.properties.spacing.after).to eq(OoxmlParser::OoxmlSize.new(2.54, :centimeter))
  end
end
