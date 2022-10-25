# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::FormTextFormat do
  let(:docxf) { OoxmlParser::Parser.parse('spec/document/elements/paragraph/std/sdt_form_properties/form.docxf') }

  it 'Has type' do
    expect(docxf.elements.first.sdt.properties.form_text_properties.format.type).to eq(:mask)
  end

  it 'Has value' do
    expect(docxf.elements.first.sdt.properties.form_text_properties.format.value).to eq('(999)999-9999')
  end

  it 'Has symbols' do
    expect(docxf.elements.first.sdt.properties.form_text_properties.format.symbols).to eq('567')
  end
end
