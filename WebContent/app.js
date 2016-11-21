var myApp = angular.module("app", []);
         
myApp.service('dataService', function($http) {
    delete $http.defaults.headers.common['X-Requested-With'];
    this.getData = function(callbackFunc) {
        $http({
            method: 'GET',
            url: '/XLSXReaderAngularJS/rest/excelService/excelData',
        }).success(function(data){
            // With the data succesfully returned, call our callback
            callbackFunc(data);
        }).error(function(){
            alert("error");
        });
     }
});

myApp.controller('MainCtrl', function($scope, dataService) {
    $scope.data = null;
    $scope.myFunction = function() {
    	dataService.getData(function(dataResponse) {
            console.log(dataResponse);
    		$scope.data = dataResponse;
        });
    }
});