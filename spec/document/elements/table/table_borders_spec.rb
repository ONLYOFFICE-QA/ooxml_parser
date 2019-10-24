# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Borders do
  it 'TableBorders' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/borders/table_borders.docx')
    expect(docx.element_by_description(location: :canvas, type: :table)[0].borders.right.val).to eq(:none)
    expect(docx.element_by_description(location: :canvas, type: :table)[1].borders.left.val).to eq(:none)
  end

  it 'TableBorderVisual' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/borders/table_borders_visual.docx')
    expect(docx.element_by_description(location: :canvas, type: :table)[0].borders.border_visual_type).to eq(:none)
  end
end
