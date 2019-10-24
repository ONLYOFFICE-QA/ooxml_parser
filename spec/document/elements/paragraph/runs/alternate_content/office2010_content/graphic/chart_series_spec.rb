# frozen_string_literal: true

require 'spec_helper'

describe 'Chart Grouping' do
  it 'do not crash chart_series_no_values' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/runs/alternate_content/office2010_content/graphic/series/chart_series_no_values.docx')
    expect(docx).to be_with_data
  end
end
