# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::FormTextComb do
  let(:docxf) { OoxmlParser::Parser.parse('spec/document/elements/paragraph/std/sdt_form_properties/form.docxf') }

  it 'Has width' do
    expect(docxf.elements.first.sdt.properties.form_text_properties.comb.width).to eq(OoxmlParser::OoxmlSize.new(77.0))
  end

  it 'Has width_rule' do
    expect(docxf.elements.first.sdt.properties.form_text_properties.comb.width_rule).to eq(:exact)
  end
end
