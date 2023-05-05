# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  docxf = OoxmlParser::Parser.parse('spec/document/elements/paragraph/std/sdt_form_properties/form_no_autofit.docxf')
  sdt_properties = docxf.elements[0].nonempty_runs[0].alternate_content.office2010_content
                        .graphic.data.text_body.elements[0].nonempty_runs[0].properties

  it 'required' do
    expect(sdt_properties.form_properties.required).to be_falsey
  end

  it 'multiline' do
    expect(sdt_properties.form_text_properties.multiline).to be_falsey
  end

  it 'autofit' do
    expect(sdt_properties.form_text_properties.autofit).to be_falsey
  end
end
