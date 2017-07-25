require 'spec_helper'

describe 'My behaviour' do
  it 'PresentationRelationships' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/relationships/presentation_relationships.pptx')
    expect(pptx.relationships[0].target).to eq('slideMasters/slideMaster1.xml')
  end
end
