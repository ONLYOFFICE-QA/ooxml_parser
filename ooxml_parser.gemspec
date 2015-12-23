Gem::Specification.new do |s|
  s.name        = 'ooxml_parser'
  s.version     = '0.1.0'
  s.date        = '2015-12-23'
  s.summary     = 'OoxmlParser Gem'
  s.description = 'Parse OOXML files (docx, xlsx, pptx)'
  s.authors     = ['Pavel Lobashov', 'Roman Zagudaev']
  s.email       = ['shockwavenn@gmail.com', 'roman.zagudaev@gmail.com']
  s.files = `git ls-files`.split($RS).reject do |file|
    file =~ %r{^(?:
    spec/.*
    |Gemfile
    |Rakefile
    |\.rspec
    |\.gitignore
    |\.rubocop.yml
    |\.rubocop_todo.yml
    |\.travis.yml
    |.*\.eps
    )$}x
  end
  s.homepage    =
      'http://rubygems.org/gems/ooxml_parser'
  s.license       = 'MIT'
end