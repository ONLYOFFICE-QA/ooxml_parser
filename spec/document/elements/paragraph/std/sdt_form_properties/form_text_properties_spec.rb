# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::FormTextProperties do
  let(:docxf) { OoxmlParser::Parser.parse('spec/document/elements/paragraph/std/sdt_form_properties/form.docxf') }

  it 'Has multiline' do
    expect(docxf.elements.first.sdt.properties.form_text_properties.multiline).to be_truthy
  end

  it 'Has autofit' do
    expect(docxf.elements.first.sdt.properties.form_text_properties.autofit).to be_truthy
  end

  it 'Has comb' do
    expect(docxf.elements.first.sdt.properties.form_text_properties.comb).to be_a(OoxmlParser::FormTextComb)
  end

  it 'Has maximum_characters' do
    expect(docxf.elements.first.sdt.properties.form_text_properties.maximum_characters.value).to eq(13)
  end

  it 'Has format' do
    expect(docxf.elements.first.sdt.properties.form_text_properties.format).to be_a(OoxmlParser::FormTextFormat)
  end
end
