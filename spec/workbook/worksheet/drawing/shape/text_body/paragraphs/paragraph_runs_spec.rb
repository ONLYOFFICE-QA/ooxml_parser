# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'ShapesRuns' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/text_body/paragraphs/runs/shape_paragraph_run.xlsx')
    expect(xlsx.worksheets[0].drawings[0].shape.text_body.paragraphs[0].runs).not_to be_empty
  end
end
