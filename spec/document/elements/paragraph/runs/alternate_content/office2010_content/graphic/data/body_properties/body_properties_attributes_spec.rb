# frozen_string_literal: true

require 'spec_helper'

describe 'body_properties_attributes' do
  it 'body_properties_text_rotation' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph'\
                                     '/runs/alternate_content/office2010_content'\
                                     '/graphic/data/body_properties/attributes'\
                                     '/body_properties_text_rotation.docx')
    expect(docx.element_by_description[0]
               .character_style_array[0].alternate_content
               .office2010_content.graphic.data
               .body_properties.vertical).to eq(:vert270)
  end
end
