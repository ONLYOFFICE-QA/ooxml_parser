# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::CheckBoxState do
  let(:docxf) do
    OoxmlParser::Parser.parse('spec/document/elements/paragraph/std/sdt_form_properties' \
                              '/checkbox/checkbox.docxf')
  end

  it 'Has value' do
    expect(docxf.elements.first.sdt.properties.checkbox.checked_state.value).to eq(25)
  end

  it 'Has font' do
    expect(docxf.elements.first.sdt.properties.checkbox.checked_state.font).to eq('Segoe UI Symbol')
  end
end
