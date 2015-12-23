require_relative 'numbering_properties/numbering_level'
module OoxmlParser
  class NumberingProperties < OOXMLDocumentObject
    attr_accessor :id, :nsid, :multilevel_type, :tmpl, :ilvls

    def initialize(id = nil, multilevel_type = nil, ilvls = [])
      @id = id
      @multilevel_type = multilevel_type
      @ilvls = ilvls
    end

    def self.parse(id)
      doc = Nokogiri::XML(File.open(OOXMLDocumentObject.path_to_folder + 'word/numbering.xml'), 'r:UTF-8')
      numbering = NumberingProperties.new
      numbering.id = id
      doc.search('//w:num').each do |num|
        next unless id == num.attribute('numId').value
        num.xpath('w:abstractNumId').each do |abstract_num_id|
          abstract_num_id = abstract_num_id.attribute('val').value
          doc.search('//w:abstractNum').each do |abstract_num|
            next unless abstract_num_id == abstract_num.attribute('abstractNumId').value
            abstract_num.xpath('w:multiLevelType').each do |multilevel_type|
              numbering.multilevel_type = multilevel_type.attribute('val').value
            end
            lvls = []
            abstract_num.xpath('w:lvl').each do |lvl|
              numbering_lvl = NumberingLevel.new
              numbering_lvl.ilvl = lvl.attribute('ilvl').value
              lvl.xpath('w:start').each do |start|
                numbering_lvl.start = start.attribute('val').value
              end
              lvl.xpath('w:numFmt').each do |num_format|
                numbering_lvl.num_format = num_format.attribute('val').value
              end
              lvl.xpath('w:lvlText').each do |level_text|
                numbering_lvl.level_text = level_text.attribute('val').value
              end
              lvl.xpath('w:lvlJc').each do |level_jc|
                numbering_lvl.level_jc = level_jc.attribute('val').value
              end
              lvl.xpath('w:pPr').each do |p_pr|
                p_pr.xpath('w:ind').each do |ind|
                  numbering_lvl.ind = Indents.parse(ind)
                end
                lvl.xpath('w:rPr').each do |r_pr|
                  r_pr.xpath('w:rFonts').each do |r_fonts|
                    if !r_fonts.attribute('ascii').nil?
                      numbering_lvl.font = r_fonts.attribute('ascii').value
                    elsif !r_fonts.attribute('hAnsi').nil?
                      numbering_lvl.font = r_fonts.attribute('hAnsi').value
                    elsif !r_fonts.attribute('eastAsia').nil?
                      numbering_lvl.font = r_fonts.attribute('eastAsia').value
                    else
                      numbering_lvl.font = r_fonts.attribute('cs').value unless r_fonts.attribute('cs').nil?
                    end
                  end
                end
              end
              lvls << numbering_lvl
            end
            numbering.ilvls = lvls
          end
        end
      end

      numbering
    end
  end
end
