module OoxmlParser
  class Note < OOXMLDocumentObject
    attr_accessor :type, :elements, :assigned_to

    def initialize(type = nil, elements = [], assigned_to = nil)
      @type = type
      @elements = elements
      @assigned_to = assigned_to
    end

    def self.parse(default_paragraph, default_character, target, assigned_to, type, parent: nil)
      note = Note.new
      note.type = type
      note.assigned_to = assigned_to
      note.parent = parent
      doc = Nokogiri::XML(File.open(OOXMLDocumentObject.path_to_folder + "word/#{target}"))
      if type.include?('footer')
        xpath_note = '//w:ftr'
      elsif type.include?('header')
        xpath_note = '//w:hdr'
      else
        raise "Cannot parse unknown Note type: #{type}"
      end
      doc.search(xpath_note).each do |ftr|
        number = 0
        ftr.xpath('*').each do |sub_element|
          if sub_element.name == 'p'
            note.elements << default_paragraph.copy.parse(sub_element, number, default_character, parent: note)
            number += 1
          elsif sub_element.name == 'tbl'
            note.elements << Table.new(parent: note).parse(sub_element, number)
            number += 1
          elsif sub_element.name == 'std'
            sub_element.xpath('w:p').each do |p|
              note.elements << default_paragraph.copy.parse(p, number, default_character)
              number += 1
            end
            sub_element.xpath('w:sdtContent').each do |sdt_content|
              sdt_content.xpath('w:p').each do |p|
                note.elements << default_paragraph.copy.parse(p, number, default_character)
                number += 1
              end
            end
          end
        end
      end
      note
    end
  end
end
