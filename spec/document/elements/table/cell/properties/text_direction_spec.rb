require 'spec_helper'

describe 'My behaviour' do
  it 'table_cell_text_no_rotation' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/cell/properties/text_direction/table_cell_text_no_rotation.docx')
    expect(docx.elements[1].rows.first.cells.first.properties.text_direction).to eq(:horizontal)
  end

  it 'table_cell_text_rotated_90' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/cell/properties/text_direction/table_cell_text_rotated_90.docx')
    expect(docx.elements[1].rows.first.cells.first.properties.text_direction).to eq(:rotate_on_90)
  end

  it 'table_cell_text_rotated_270' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/cell/properties/text_direction/table_cell_text_rotated_270.docx')
    expect(docx.elements[1].rows.first.cells.first.properties.text_direction).to eq(:rotate_on_270)
  end

  it 'table_cell_text_left_to_right_top_to_bottom' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/cell/properties/text_direction/table_cell_text_left_to_right_top_to_bottom.docx')
    expect(docx.elements[1].rows.first.cells.first.properties.text_direction).to eq(:left_to_right_top_to_bottom)
  end

  it 'table_cell_text_top_to_bottom_right_to_left' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/cell/properties/text_direction/table_cell_text_top_to_bottom_right_to_left.docx')
    expect(docx.elements[1].rows[2].cells[1].properties.text_direction).to eq(:top_to_bottom_right_to_left)
  end

  it 'table_cell_text_bottom_to_top_left_to_right' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/cell/properties/text_direction/table_cell_text_bottom_to_top_left_to_right.docx')
    expect(docx.elements[1].rows[2].cells[2].properties.text_direction).to eq(:bottom_to_top_left_to_right)
  end
end
