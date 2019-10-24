# frozen_string_literal: true

require 'spec_helper'

describe 'VerticalMerge' do
  it 'cell_vertical_merge_value' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/cell/properties/vertical_merge/cell_vertical_merge_value.docx')
    expect(docx.elements[1].rows[1].cells[1].properties.vertical_merge.value).to eq(:restart)
  end
end
