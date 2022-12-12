import copy
import argparse
import sys
import xlrd
from decimal import *
from copy import copy
from random import randrange
from itertools import permutations


class Ambulance():
    def __init__(self, name, cap, hopital, greenCase, redCase):
        self.name = name
        self.cap = cap
        self.hopital = hopital
        self.liveCapacity = cap
        self.start = hopital
        self.time = 0
        self.routes = []
        self.patientAvailableToCover = []
        self.route_possibilities = []
        self.patientType = 'green'
        if redCase == 1:
            self.patientType = 'red'

    def __str__(self):
        str = f"Ambulance: {self.name}\n  Capacity: {self.cap}\n  live Capacity: {self.liveCapacity}\n  Hopital: {self.hopital}\n  Patient Type: {self.patientType}\n  start: {self.start} \n  patientAvailableToCover :"
        str += " ".join([t.name for t in self.patientAvailableToCover])  
        return str 


class Job():
    def __init__(self, name, priority, duration, coveredBy):
        self.name = name
        self.priority = priority
        self.duration = duration
        self.coveredBy = coveredBy

    def __str__(self): 
        about = f"Job: {self.name}\n  Priority: {self.priority}\n  Duration: {self.duration}\n  Covered by: "
        about += ", ".join([t.name for t in self.coveredBy])
        return about
		
class Patient():
    def __init__(self, name, loc, job, tStart, tEnd, tDue):
        self.name = name
        self.loc = loc
        self.job = job
        self.tStart = tStart
        self.tEnd = tEnd
        self.tDue = tDue
        self.served = 0

    def __str__(self):
        coveredBy = ", ".join([t.name for t in self.job.coveredBy])
        return f"Patient: {self.name}\n  Location: {self.loc}\n  Job: {self.job.name}\n  Priority: {self.job.priority}\n  Duration: {self.job.duration}\n  Covered by: {coveredBy}\n  Start time: {self.tStart}\n  End time: {self.tEnd}\n  Due time: {self.tDue}\n  Covered: {self.served}"

class Route():
    def __init__(self, from_node, to_node, dist, t, proc, tLate, patient_priority, is_start, is_end):
        self.from_node = from_node
        self.to_node = to_node
        self.dist = dist
        self.t = t
        self.tLate = tLate
        self.patient_priority = patient_priority
        self.proc = proc
        self.is_start = is_start
        self.is_end = is_end
        self.covered = 0

    def __str__(self):
        if self.is_start == 1:
            return f"{self.from_node}"
        if self.is_end == 1:
            return f" -> {self.to_node} (dist={self.dist})" 
        return f" -> {self.to_node} (dist={self.dist}, t={self.t}, proc={self.proc}, tLate={self.tLate})"

		# return f"from_node: {self.from_node}\n  to_node: {self.to_node}\n  dist: {self.dist}\n  t: {self.t}\n  proc: {self.proc}\n  is_start: {self.is_start}\n  is_end: {self.is_end}"

class Solution():
	def __init__(self, name, total_late_time, assignedStr, str_capacity_utilization, str_all_route, total_live_capacity, total_capacity, total_percent, routes):
		self.name = name
		self.total_late_time = total_late_time
		self.assignedStr = assignedStr
		self.str_capacity_utilization = str_capacity_utilization
		self.str_all_route = str_all_route
		self.total_live_capacity = total_live_capacity
		self.total_capacity = total_capacity
		self.total_percent = total_percent
		self.routes = routes

	def __str__(self):
		result = "\n*********************** " + self.name + " *********************** \n"
		result += "\ntotal waiting time : " + str(round(self.total_late_time, 2)) + " minutes\n"
		result += self.assignedStr + "\n"
		result += self.str_all_route + "\n"
		result += self.str_capacity_utilization + "\n"
		result += "\nTotal Ambulance utilization is " + str(self.total_percent) + "% (" + str(self.total_live_capacity) + "/" + str(self.total_capacity) + ")"
		result += "\n*************************************************** \n"
		return result

def create_data_model(filename):
	wb = xlrd.open_workbook(filename)

	# Read Ambulance data
	ws = wb.sheet_by_name('Ambulances')
	ambulances = []
	for i,t in enumerate(ws.col_values(0)[3:]):
		# Create Ambulance object
		thisTech = Ambulance(*ws.row_values(3+i)[:5])
		ambulances.append(thisTech)

	# Read job data
	jobs = []
	for j,b in enumerate(ws.row_values(0)[3:]):
		coveredBy = [t for i,t in enumerate(ambulances) if ws.cell_value(3+i,3+j) == 1]
		# Create Job object
		thisJob = Job(*ws.col_values(3+j)[:3], coveredBy)
		jobs.append(thisJob)

	# Read location data
	ws = wb.sheet_by_name('Locations')
	locations = ws.col_values(0)[1:]
	dist = {(l, l) : 0 for l in locations}
	for i,l1 in enumerate(locations):
		for j,l2 in enumerate(locations):
			if i < j:
				dist[l1,l2] = ws.cell_value(1+i, 1+j)
				dist[l2,l1] = dist[l1,l2]

	# Read patient data
	ws = wb.sheet_by_name('Patients')
	patients = []
	for i,c in enumerate(ws.col_values(0)[1:]):
		for b in jobs:
			if b.name == ws.cell_value(1+i, 2):
				# Create Patient object using corresponding Job object
				rowVals = ws.row_values(1+i)
				thisPatient = Patient(*rowVals[:2], b, *rowVals[3:])
				patients.append(thisPatient)
				for ambulance in b.coveredBy:
					ambulance.patientAvailableToCover.append(thisPatient)

				break
	return ambulances, jobs, dist, patients

def get_closest_patient_to_served(ambulance, dist, tabu_search = False):
	index_patient = -1
	min_dist = -1
	index = 0
	min_tStart = 0
	min_tEnd = 0
	index_patients_to_covered = []

	min_dist_patients_to_covered = []

	index = 0
	index_patient = -1
	for patient in ambulance.patientAvailableToCover:

		if patient.served == 0:
			if (min_dist == -1 or dist[ambulance.start, patient.loc] < min_dist) and patient.served == 0:
				duration = dist[ambulance.start, patient.loc] + patient.job.duration + dist[patient.loc, ambulance.hopital]
				if ambulance.liveCapacity - duration >= 0:
					min_dist = dist[ambulance.start, patient.loc]
					min_tStart = patient.tStart
					min_tEnd = patient.tEnd
					index_patient = index
					index_patients_to_covered.append(index)
					min_dist_patients_to_covered.append(min_dist)
			elif (min_dist == -1 or dist[ambulance.start, patient.loc] == min_dist) and patient.served == 0:
				if patient.tStart < min_tStart:
					duration = dist[ambulance.start, patient.loc] + patient.job.duration + dist[patient.loc, ambulance.hopital]
					if ambulance.liveCapacity - duration >= 0:
						min_dist = dist[ambulance.start, patient.loc]
						min_tStart = patient.tStart
						min_tEnd = patient.tEnd
						index_patient = index
						index_patients_to_covered.append(index)
						min_dist_patients_to_covered.append(min_dist)

				elif patient.tStart == min_tStart:
					if patient.tEnd <= min_tEnd:
						duration = dist[ambulance.start, patient.loc] + patient.job.duration + dist[patient.loc, ambulance.hopital]
						if ambulance.liveCapacity - duration >= 0:
							min_dist = dist[ambulance.start, patient.loc]
							min_tStart = patient.tStart
							min_tEnd = patient.tEnd
							index_patient = index
							index_patients_to_covered.append(index)
							min_dist_patients_to_covered.append(min_dist)

		index += 1
	if tabu_search == True:
		if len(index_patients_to_covered) > 1:
			r = randrange(len(index_patients_to_covered) - 1)
			while r == index_patient:
				r = randrange(len(index_patients_to_covered))
			if r != index_patient:
				return index_patients_to_covered[r], min_dist_patients_to_covered[r]

				
	return index_patient, min_dist
	
def assign_ambulance_to_closest_patient_not_served(ambulances, dist, patient, m, tabu_search = False, first_solution = True, i=0):
    total_late_time = 0
    jobStr = ''
    index_ambulance = -1
    min_dist = 0
    index = 0

    for ambulance in patient.job.coveredBy:
        if (min_dist == 0 or dist[ambulance.start, patient.loc] < min_dist):
            min_dist = dist[ambulance.start, patient.loc]
            index_ambulance = index
        index += 1

    if index_ambulance == -1:
        return '', total_late_time

    else:
    
        ambulance = patient.job.coveredBy[index_ambulance]
        if patient.served == 0:
            start = ambulance.start
            duration = dist[start, patient.loc] + patient.job.duration

            patient.served = 1
            ambulance.liveCapacity -= duration
            ambulance.start = patient.loc
            jobStr += "\n{} assigned to {} ({}) in {}. Start at t={:.2f}.".format(ambulance.name,patient.name,patient.job.name,patient.loc,ambulance.time + dist[start, patient.loc])

            late_time = ambulance.time + dist[start, patient.loc] + patient.job.duration - patient.tDue
            if late_time > 0:
                jobStr += " {:.2f} minutes late.".format(late_time)
                total_late_time += late_time * patient.job.priority
            else:
                late_time = 0
            if patient.tStart > ambulance.time:
                jobStr += " Start time corrected by {:.2f} minutes.".format(patient.tStart - ambulance.time )
            if ambulance.time > patient.tEnd:
                jobStr += " End time corrected by {:.2f} minutes.".format(ambulance.time - patient.tEnd)
				
            is_start = 0
            is_end = 0
            if start == ambulance.hopital:
                is_start = 1
                route = Route(start, patient.loc, dist[start, patient.loc], ambulance.time, patient.job.duration, late_time, patient.job.priority, is_start, is_end)
                ambulance.routes.append(route)
            route = Route(start, patient.loc, dist[start, patient.loc], ambulance.time + dist[start, patient.loc], patient.job.duration, late_time, patient.job.priority, 0, 0)
            ambulance.routes.append(route)
			
            ambulance.time += dist[start, patient.loc] + patient.job.duration		
	
    return jobStr, total_late_time, tabu_search

def generate_solution(file, m, tabu_search = False, first_solution = True):
 
	# print(tabu_search)
	ambulances, jobs, dist, patients = create_data_model(file)

	assignedStr = ''
	total_late_time = 0
	i = 0
	for patient in patients:
		jobStr, late_time, tabu_search = assign_ambulance_to_closest_patient_not_served(ambulances, dist, patient, m, tabu_search, first_solution, i)
		i+=1
		if randrange(10) == 3:
			if tabu_search == True:
				tabu_search = False
		assignedStr += jobStr
		total_late_time += late_time

	for patient in patients:
		if patient.served == 0:
		    assignedStr += '\nNobody assigned to ' + patient.name +' ('+ patient.job.name+') in '+ patient.loc
		    total_late_time += patient.job.priority * m

	total_percent = 0
	str_capacity_utilization = ''
	total_live_capacity = 0
	total_capacity = 0
	str_all_route = ''
	for ambulance in ambulances:
		if (len(ambulance.routes) > 0):
			route = Route(ambulance.start, ambulance.hopital, dist[ambulance.start, ambulance.hopital], ambulance.time, 0, 0, 0, 0, 1)
			ambulance.routes.append(route)
		
		ambulance.liveCapacity -= dist[ambulance.start, ambulance.hopital]
		ambulance.time += dist[ambulance.start, ambulance.hopital]
		ambulance.start = ambulance.hopital
		str_route = ambulance.name + "'s route: "
		for route in ambulance.routes:
		    str_route += str(route)
		total_capacity += ambulance.cap
		total_live_capacity += ambulance.cap - ambulance.liveCapacity
		percent = round(((ambulance.cap - ambulance.liveCapacity) * 100) / ambulance.cap, 2)

		str_capacity_utilization += "\n" + ambulance.name + "'s utilization is " + str(percent) + "% (" + str(ambulance.cap - ambulance.liveCapacity) + "/" + str(ambulance.cap) + ")"
		str_all_route += "\n" + str_route
	total_percent = round((total_live_capacity * 100) / total_capacity, 2)	
	
	solution = Solution("First solution", total_late_time, assignedStr, str_capacity_utilization, str_all_route, total_live_capacity, total_capacity, total_percent, ambulances)

	for ambulance in solution.routes: 
		array_patient = []
		index = 0
		for route in ambulance.routes: 
			if index > 0 and index < len(ambulance.routes) - 1:
				array_patient.append(route.to_node)
			index += 1
		ambulance.route_possibilities = list(permutations(array_patient))

	return solution, tabu_search

def compare_solution(best_solution, next_solution):
    """ compare tow solutions 
	return 1 if next_solution.total_late_time < next_solution.total_late_time
	"""
    if next_solution.total_late_time < best_solution.total_late_time:
        return 1
    return 0

def get_next_neighborhood(solution, count):
	tmp_neighborhood = None
	i = 0
	index = 0
	route_index = 0
	count_all_possibility = 0
	for route in solution.routes:
		count_all_possibility += len(route.route_possibilities)
	for route in solution.routes:
		for new_route in route.route_possibilities:
			if i == count:
				tmp_neighborhood = new_route[:]
				route_index = index
			i += 1
			# if i >= count_all_possibility:
				# i = 0
		index += 1

	if tmp_neighborhood:
		j = 0
		for route in solution.routes[route_index].routes:
			if j > 0 and j < len(solution.routes[route_index].routes) -1:
				route.to_node = tmp_neighborhood[j - 1]
			j += 1

	return solution, count_all_possibility

def two_opt(solution):
	num_ambulance = randrange(len(solution.routes))
	
	count = 0
	while len(solution.routes[num_ambulance].routes) < 4 or count < 100:
		num_ambulance = randrange(0, len(solution.routes))
		count += 1
		
	if len(solution.routes[num_ambulance].routes) >= 4:
		index_route = randrange(1, len(solution.routes[num_ambulance].routes) - 2)
		tmp = solution.routes[num_ambulance].routes[index_route + 1].to_node
		solution.routes[num_ambulance].routes[index_route + 1].to_node = solution.routes[num_ambulance].routes[index_route].to_node
		solution.routes[num_ambulance].routes[index_route].to_node = tmp
	
	return solution

def or_opt(solution):
	num_ambulance = randrange(len(solution.routes))
	
	count = 0
	while len(solution.routes[num_ambulance].routes) < 5 or count < 100:
		num_ambulance = randrange(0, len(solution.routes))
		count += 1
		
	if len(solution.routes[num_ambulance].routes) >= 5:
		index_route1 = randrange(1, len(solution.routes[num_ambulance].routes) - 2)
		index_route2 = randrange(1, len(solution.routes[num_ambulance].routes) - 2)
		while index_route2 == index_route1:
			index_route2 = randrange(1, len(solution.routes[num_ambulance].routes) - 2)
		
		change_index_1_to_route = randrange(1, len(solution.routes[num_ambulance].routes) - 2)
		while change_index_1_to_route == index_route1:
			change_index_1_to_route = randrange(1, len(solution.routes[num_ambulance].routes) - 2)			

		change_index_2_to_route = randrange(1, len(solution.routes[num_ambulance].routes) - 2)
		while change_index_2_to_route == index_route2:
			change_index_2_to_route = randrange(1, len(solution.routes[num_ambulance].routes) - 2)		
		
		tmp1 = solution.routes[num_ambulance].routes[change_index_1_to_route].to_node
		solution.routes[num_ambulance].routes[change_index_1_to_route].to_node = solution.routes[num_ambulance].routes[index_route1].to_node
		solution.routes[num_ambulance].routes[index_route1].to_node = tmp1		

		tmp2 = solution.routes[num_ambulance].routes[change_index_2_to_route].to_node
		solution.routes[num_ambulance].routes[change_index_2_to_route].to_node = solution.routes[num_ambulance].routes[index_route1].to_node
		solution.routes[num_ambulance].routes[index_route1].to_node = tmp2
	
	return solution


def shift_opt(solution, patients):
	num_ambulance = 0
	min_tLate = 0
	min_utilization = -1
	from_ambulance = -1
	to_ambulance = -1
	from_route = 0
	to_route = 0
	index = 0

	for ambulance in solution.routes:
		index2 = 0
		for route in ambulance.routes:
			if route.tLate > min_tLate:
				from_ambulance = index
				min_tLate = route.tLate
				from_route = index2
			index2 += 1
		index += 1		


	patient = get_patient_by_loc(patients, solution.routes[from_ambulance].routes[from_route].from_node)
	index = 0	
	for ambulance in solution.routes:
		can_assign = False
		for k in patient.job.coveredBy:
			if k.name == ambulance.name:
				can_assign = True
		if can_assign:
			if ambulance.cap - ambulance.liveCapacity > min_utilization :
				to_ambulance = index
				min_utilization = ambulance.cap - ambulance.liveCapacity

		index += 1	

	print(to_ambulance)
	exit()
	if from_ambulance != -1 and to_ambulance != -1 and to_ambulance != from_ambulance:
		solution.routes[to_ambulance].routes.append(solution.routes[from_ambulance].routes[from_route])
		solution.routes[from_ambulance].routes.remove(solution.routes[from_ambulance].routes[from_route])
		print("iciiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiii")
		exit()
		print("to_ambulance = " )
		print( to_ambulance )
		print("from_route = ")
		print( from_route)
		print("from_ambulance = " )
		print(from_ambulance)
		print("---")
	
	return solution


def get_patient_by_loc(patients, loc):
	for patient in patients:
		if patient.loc == loc:
			return patient
	return None

def get_ambulance_by_name(ambulances, name):
	for ambulance in ambulances:
		if ambulance.name == name:
			return ambulance
	return None

def optimize_solution(file, solution, m, count):
    ambulances, jobs, dist, patients = create_data_model(file)
    tmp_solution = copy(solution)
    tmp_solution, count_all_possibility = get_next_neighborhood(tmp_solution, count)
# get random option 
    total_late_time = 0
    jobStr = ''

    if count > count_all_possibility:
        option = randrange(2)
        if option == 0:
            tmp_solution = two_opt(tmp_solution)
        if option == 1:
            tmp_solution = or_opt(tmp_solution)
		# if option == 2:
			# tmp_solution = shift_opt(tmp_solution, patients)

    for tmp_ambulance in tmp_solution.routes:
        i = 0
        for route in tmp_ambulance.routes:
            if i > 0 and i < len(tmp_ambulance.routes) - 1:
                ambulance = get_ambulance_by_name(ambulances, tmp_ambulance.name)
                patient = get_patient_by_loc(patients, route.to_node)
                if patient:
# start
                    start = ambulance.start
                    duration = dist[start, patient.loc] + patient.job.duration

                    patient.served = 1
                    ambulance.liveCapacity -= duration
                    ambulance.start = patient.loc
                    jobStr += "\n{} assigned to {} ({}) in {}. Start at t={:.2f}.".format(ambulance.name,patient.name,patient.job.name,patient.loc,ambulance.time + dist[start, patient.loc])

                    late_time = ambulance.time + dist[start, patient.loc] + patient.job.duration - patient.tDue
                    if late_time > 0:
                        jobStr += " {:.2f} minutes late.".format(late_time)
                        total_late_time += late_time * patient.job.priority
                    else:
                        late_time = 0
                    if patient.tStart > ambulance.time:
                        jobStr += " Start time corrected by {:.2f} minutes.".format(patient.tStart - ambulance.time )
                    if ambulance.time > patient.tEnd:
                        jobStr += " End time corrected by {:.2f} minutes.".format(ambulance.time - patient.tEnd)
						
                    is_start = 0
                    is_end = 0
                    if start == ambulance.hopital:
                        is_start = 1
                        route = Route(start, patient.loc, dist[start, patient.loc], ambulance.time, patient.job.duration, late_time, patient.job.priority, is_start, is_end)
                        ambulance.routes.append(route)
                    route = Route(start, patient.loc, dist[start, patient.loc], ambulance.time + dist[start, patient.loc], patient.job.duration, late_time, patient.job.priority, 0, 0)
                    ambulance.routes.append(route)
					
                    ambulance.time += dist[start, patient.loc] + patient.job.duration

            i += 1
# end
    assignedStr = jobStr
    for patient in patients:
        if patient.served == 0:
            assignedStr += '\nNobody assigned to ' + patient.name +' ('+ patient.job.name+') in '+ patient.loc
            total_late_time += patient.job.priority * m

    total_percent = 0
    str_capacity_utilization = ''
    total_live_capacity = 0
    total_capacity = 0
    str_all_route = ''
    for ambulance in ambulances:
        if (len(ambulance.routes) > 0):
            route = Route(ambulance.start, ambulance.hopital, dist[ambulance.start, ambulance.hopital], ambulance.time, 0, 0, 0, 0, 1)
            ambulance.routes.append(route)

        ambulance.liveCapacity -= dist[ambulance.start, ambulance.hopital]
        ambulance.time += dist[ambulance.start, ambulance.hopital]
        ambulance.start = ambulance.hopital
        str_route = ambulance.name + "'s route: "
        for route in ambulance.routes:
            str_route += str(route)
        total_capacity += ambulance.cap
        total_live_capacity += ambulance.cap - ambulance.liveCapacity
        percent = round(((ambulance.cap - ambulance.liveCapacity) * 100) / ambulance.cap, 2)

        str_capacity_utilization += "\n" + ambulance.name + "'s utilization is " + str(percent) + "% (" + str(ambulance.cap - ambulance.liveCapacity) + "/" + str(ambulance.cap) + ")"
        str_all_route += "\n" + str_route
    total_percent = round((total_live_capacity * 100) / total_capacity, 2)	
	
    solution = Solution("Solution " + str(count), total_late_time, assignedStr, str_capacity_utilization, str_all_route, total_live_capacity, total_capacity, total_percent, ambulances)
    return solution

def tabu_search(first_solution, file, m, iters, size):
    count = 1
    tabu_list = list()
    best_solution = copy(first_solution)
    all_solution = []
    tabu_search = False
    while count <= iters:
        next_solution = optimize_solution(file, best_solution, m, count)
        print(next_solution)
        # exit()
        
        if randrange(iters) == count:
            if tabu_search == True:
                tabu_search = False
        all_solution.append(next_solution)
        if compare_solution(best_solution, next_solution) == 1:
            best_solution = copy(next_solution)
       
        if len(tabu_list) >= size:
            tabu_list.pop(0)
        count += 1
    best_solution.name = "Best solution"
    return best_solution	
  

def main(args=None):
    m = 600
    # ambulances, jobs, dist, patients = create_data_model(args.File)
    # print (list(permutations(tab)))
    
    first_solution, ts = generate_solution(args.File, m, False, True)
    print(first_solution)

    best_solution = tabu_search(first_solution, args.File, m, args.Iterations, args.Size)
    print("Best solution: ")
    print(best_solution)
    
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Tabu Search")
   
    parser.add_argument(
        "-f", "--File", type=str, help="Path to the file containing the data", required=True,
    )
	
    parser.add_argument(
        "-i", "--Iterations", type=int, help="How many iterations the algorithm should perform", required=True,
    )
    parser.add_argument(
        "-s", "--Size", type=int, help="Size of the tabu list", required=True
    )

    # Pass the arguments to main method
    sys.exit(main(parser.parse_args()))
