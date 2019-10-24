# frozen_string_literal: true

require 'spec_helper'

describe 'VerticalMerge' do
  it 'cell_vertical_merge_value' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/cell/elements/table_cell_sdt_element.docx')
    expect(docx.elements[1].rows.first.cells.first.elements.first).to be_a(OoxmlParser::StructuredDocumentTag)
  end
end
