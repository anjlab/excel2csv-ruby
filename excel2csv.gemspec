# -*- encoding: utf-8 -*-
require File.expand_path('../lib/excel2csv/version', __FILE__)

Gem::Specification.new do |gem|
  gem.authors       = ["Yury Korolev"]
  gem.email         = ["yury.korolev@gmail.com"]
  gem.description   = %q{extract excel worksheets to csv files}
  gem.summary       = %q{extract excel worksheets to csv files}
  gem.homepage      = ""

  gem.executables   = `git ls-files -- bin/*`.split("\n").map{ |f| File.basename(f) }
  gem.files         = `git ls-files`.split("\n")
  gem.files.reject! { |fn| fn.include? "vendor/" }  
  gem.test_files    = `git ls-files -- {test,spec,features}/*`.split("\n")
  gem.name          = "excel2csv"
  gem.require_paths = ["lib"]
  gem.version       = Excel2CSV::VERSION

  gem.add_development_dependency "rspec", "<= 2.6"
end
