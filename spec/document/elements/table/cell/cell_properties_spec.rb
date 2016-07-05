require 'spec_helper'

describe OoxmlParser::CellProperties do
  it 'cell_properties_merge_restart' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/cell/properties/cell_properties_merge_restart.docx')
    expect(docx.elements[556].rows.first.cells.first.properties.merge.value).to eq(:restart)
  end

  it 'shade_nil' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/cell/properties/shade_nil.docx')
    expect(docx.elements.first.rows.first.cells.first.properties.shd).to eq(:none)
  end

  it 'table_cell_text_no_rotation' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/cell/properties/table_cell_text_no_rotation.docx')
    expect(docx.elements[1].rows.first.cells.first.properties.text_direction).to eq(:horizontal)
  end

  it 'table_cell_text_rotated_90' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/cell/properties/table_cell_text_rotated_90.docx')
    expect(docx.elements[1].rows.first.cells.first.properties.text_direction).to eq(:rotate_on_90)
  end

  it 'table_cell_text_rotated_270' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/cell/properties/table_cell_text_rotated_270.docx')
    expect(docx.elements[1].rows.first.cells.first.properties.text_direction).to eq(:rotate_on_270)
  end

  it 'table_custom_cell_margin.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/cell/properties/table_custom_cell_margin.docx')
    margins_in_doc = docx.element_by_description(location: :canvas, type: :paragraph)[1].rows[0].cells[0].properties.table_cell_margin
    expect(OoxmlParser::TableMargins.new(false, 0.5, 3.19, 6.0, 3.19)).to eq(margins_in_doc)
  end

  it 'table_cell_shade.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/cell/properties/table_cell_shade.docx')
    docx.element_by_description(location: :canvas, type: :paragraph)[1].rows.each do |each_row|
      each_row.cells.each do |current_cell|
        expect(current_cell.cell_properties.shd).to eq(OoxmlParser::Color.new(217, 226, 242))
      end
    end
  end

  it 'borders_properties_compare' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/cell/properties/borders_properties.docx')
    docx.element_by_description(location: :canvas, type: :paragraph)[1].rows.each do |each_row|
      (2...6).each do |column_counter|
        expect(each_row.cells[column_counter].cell_properties.borders_properties).to eq(OoxmlParser::Borders.new)
      end
    end
  end
end
