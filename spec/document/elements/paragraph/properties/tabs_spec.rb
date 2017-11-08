require 'spec_helper'

describe 'My behaviour' do
  it 'tabs_position' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/properties/tabs/tabs_position.docx')
    expect(docx.elements.first.paragraph_properties.tabs[0].position).to eq(OoxmlParser::OoxmlSize.new(1440))
    expect(docx.elements.first.paragraph_properties.tabs[0].value).to eq(:left)
  end

  it 'tab_leader' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/properties/tabs/tab_leader.docx')
    expect(docx.elements.first.paragraph_properties.tabs[0].leader).to eq(:dot)
  end
end
