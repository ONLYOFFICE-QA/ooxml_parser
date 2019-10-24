# frozen_string_literal: true

require 'spec_helper'

describe 'DrawingPropertiesLineDash' do
  it 'shape_insert_all_fill_picture.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/line/drawing_properties_line_dash/drawing_properties_line_dash_type.docx')
    alternate_content = docx.elements.first.nonempty_runs.first.alternate_content
    expect(alternate_content.office2010_content.graphic.data.properties.line.dash.value).to eq(:large_dash_dot)
  end
end
