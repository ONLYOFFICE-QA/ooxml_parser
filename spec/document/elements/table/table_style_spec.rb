require 'spec_helper'

describe OoxmlParser::TableStyle do
  it 'Check Parsing Table Style Shade None' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/style/table_style_shade_none.docx')
    expect(docx.elements[1].table_properties.table_style.first_column.cell_properties.shade).to be_nil
  end

  it 'Check Parsing Table Style Banding2 without color' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/style/table_style_columns_none.docx')
    expect(docx.elements[1].table_properties.table_style.banding_2_horizontal.cell_properties.shade).to be_nil
  end

  it 'Check Parsing Table Style Banding2 with color' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/style/table_style_banding_2.docx')
    expect(docx.elements[1].table_properties.table_style.banding_2_horizontal.cell_properties.shade.fill.to_s).to eq('RGB (217, 217, 217)')
  end

  it 'table_style_banging_2_empty' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/style/table_style_banging_2_empty.docx')
    table_style = docx.elements[1].table_properties.table_style
    expect(table_style.banding_2_horizontal.cell_properties.shade).to be_nil
  end
end
