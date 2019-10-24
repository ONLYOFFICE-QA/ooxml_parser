# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::TableProperties do
  it 'TableLook' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/properties/table_look.docx')
    table_look = docx.elements[1].table_properties.table_look
    expect(table_look.first_column).to eq(false)
    expect(table_look.first_row).to eq(false)
    expect(table_look.last_column).to eq(false)
    expect(table_look.last_row).to eq(false)
    expect(table_look.no_horizontal_banding).to eq(true)
    expect(table_look.no_vertical_banding).to eq(true)
  end

  it 'table_properties_position_x_nil' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/properties/table_properties_position_x_nil.docx')
    expect(docx.elements[1].table_properties.table_properties.position_x).to be_nil
  end

  it 'table_properties_position_x' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/properties/table_properties_position_x.docx')
    expect(docx.elements[1].table_properties.table_properties.position_x).to eq(OoxmlParser::OoxmlSize.new(2.2, :centimeter))
  end

  it 'table_cell_margin' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/properties/table_cell_margin.docx')
    margins_in_doc = docx.element_by_description(location: :canvas, type: :paragraph)[1].table_properties.table_cell_margin
    expect(OoxmlParser::TableMargins.new(true,
                                         OoxmlParser::OoxmlSize.new(1.0, :centimeter),
                                         OoxmlParser::OoxmlSize.new(0.19, :centimeter),
                                         OoxmlParser::OoxmlSize.new(2.0, :centimeter),
                                         OoxmlParser::OoxmlSize.new(0.19, :centimeter))).to eq(margins_in_doc)
  end

  describe 'JC' do
    it 'table_properties_jc_center' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/properties/table_properties_jc_center.docx')
      expect(docx.element_by_description(location: :canvas, type: :paragraph)[1]
                 .table_properties.jc).to eq(:center)
    end

    it 'table_properties_jc_left' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/properties/table_properties_jc_left.docx')
      expect(docx.element_by_description[1].table_properties.jc).to eq(:left)
    end
  end

  it 'table_shade_color' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/properties/table_shade_color.docx')
    expect(docx.element_by_description(location: :canvas, type: :paragraph)[1]
               .table_properties.shade.fill).to eq(OoxmlParser::Color.new(30, 79, 121))
  end

  it 'table_border_size' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/properties/table_border_size.docx')
    expect(docx.elements[1].table_properties.table_borders.top.sz).to eq(OoxmlParser::OoxmlSize.new(4.5, :point))
  end
end
