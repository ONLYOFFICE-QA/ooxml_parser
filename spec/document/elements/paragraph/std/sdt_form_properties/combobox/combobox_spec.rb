# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::ComboBox do
  let(:docx) do
    OoxmlParser::Parser.parse('spec/document/elements/paragraph/std/sdt_form_properties' \
                              '/combobox/combobox.docx')
  end

  it 'Item has text' do
    expect(docx.elements.first.sdt.properties.combobox.list_items[0].text).to eq('1')
  end

  it 'Item has value' do
    expect(docx.elements.first.sdt.properties.combobox.list_items[0].value).to eq('2')
  end
end
