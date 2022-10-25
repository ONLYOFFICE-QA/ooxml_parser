# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::FormProperties do
  let(:docxf) { OoxmlParser::Parser.parse('spec/document/elements/paragraph/std/sdt_form_properties/form.docxf') }

  it 'Has key' do
    expect(docxf.elements.first.sdt.properties.form_properties.key).to eq('123')
  end

  it 'Has help_text' do
    expect(docxf.elements.first.sdt.properties.form_properties.help_text).to eq('Phone number')
  end

  it 'Has required' do
    expect(docxf.elements.first.sdt.properties.form_properties.required).to be_truthy
  end

  it 'Has shade' do
    expect(docxf.elements.first.sdt.properties.form_properties.shade.fill).to eq(OoxmlParser::Color.new(244, 176, 131))
  end

  it 'Has border' do
    expect(docxf.elements.first.sdt.properties.form_properties.border.color).to eq(OoxmlParser::Color.new(102, 102, 153))
  end
end
