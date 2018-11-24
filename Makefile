.PHONY: build deploy

build:  
    npm install --silent
    tsc --pretty

deploy: build  
    clasp push