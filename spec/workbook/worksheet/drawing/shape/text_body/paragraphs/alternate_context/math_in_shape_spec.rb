require 'spec_helper'

describe 'My behaviour' do
  it 'math_in_shape' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/text_body/paragraphs/alternate_context/math/math_in_shape.xlsx')
    expect(xlsx.worksheets[0].drawings[0].shape.text_body.paragraphs[0].alternate_content
          .choice.math_text.math_paragraph.math.formula_run.first.text).not_to be_empty
  end
end
