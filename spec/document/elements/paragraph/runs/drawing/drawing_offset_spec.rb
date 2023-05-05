# frozen_string_literal: true

require 'spec_helper'

describe 'Drawing Offset' do
  folder = 'spec/document/elements/paragraph/runs/drawing/drawing_offset'

  it 'chart_offset_w14' do
    docx = OoxmlParser::DocxParser.parse_docx("#{folder}/chart_offset_w14.docx")
    elements = docx.element_by_description(location: :canvas, type: :paragraph)
    expect(elements.first.nonempty_runs.first.drawing.properties.horizontal_position.offset).to be_zero
    expect(elements.first.nonempty_runs.first.drawing.properties.vertical_position.offset).to be_zero
  end

  it 'Check Diagram offset should !=0' do
    docx = OoxmlParser::DocxParser.parse_docx("#{folder}/chart_offset_non_zero.docx")
    expect(docx.elements.first.character_style_array.first.drawing.properties.horizontal_position.offset).not_to be_zero
    expect(docx.elements.first.character_style_array.first.drawing.properties.vertical_position.offset).not_to be_zero
  end

  it 'Check Diagram offset should ==0' do
    docx = OoxmlParser::DocxParser.parse_docx("#{folder}/chart_offset_zero.docx")
    expect(docx.elements.first.character_style_array.first.drawing.properties.horizontal_position.offset).to be_zero
    expect(docx.elements.first.character_style_array.first.drawing.properties.vertical_position.offset).to be_zero
  end
end
