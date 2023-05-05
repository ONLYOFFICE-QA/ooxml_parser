# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::CheckBox do
  docxf = OoxmlParser::Parser.parse('spec/document/elements/paragraph/std/sdt_form_properties' \
                                    '/checkbox/checkbox.docxf')

  it 'Has checked' do
    expect(docxf.elements.first.sdt.properties.checkbox.checked).to be_falsey
  end

  it 'Has checked_state' do
    expect(docxf.elements.first.sdt.properties.checkbox.checked_state).to be_a(OoxmlParser::CheckBoxState)
  end

  it 'Has unchecked_state' do
    expect(docxf.elements.first.sdt.properties.checkbox.unchecked_state).to be_a(OoxmlParser::CheckBoxState)
  end

  it 'Has group_key' do
    expect(docxf.elements.first.sdt.properties.checkbox.group_key.value).to eq('Group 1')
  end
end
